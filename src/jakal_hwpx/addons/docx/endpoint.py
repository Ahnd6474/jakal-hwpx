"""Versioned Word editing service and an optional standard-library HTTP adapter."""
from __future__ import annotations

from hashlib import sha256
from collections import OrderedDict
from copy import deepcopy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from threading import RLock
from urllib.parse import urlsplit, parse_qs

from lxml import etree

from .editing import WordContainer, WordParagraph, WordTable, WordPicture, _tag
from .identity import ID, STORIES, stamp, identify


class VersionConflict(ValueError):
    """The caller's revision no longer describes the document."""


class ObjectNotFound(KeyError):
    """A persistent object identity is absent from this revision."""


def _owner(editor, node, part):
    from docx.blkcntnr import BlockItemContainer
    roots = {_tag(tag) for tag in ("body", "tc", "txbxContent", "hdr", "ftr", "footnote", "endnote")}
    root = next((parent for parent in node.iterancestors() if parent.tag in roots), None)
    if root is None:
        root = editor._backend.element.body
        part = editor._backend.part
    return WordContainer(editor, root, BlockItemContainer(root, part))


def resolve_object(editor, object_id):
    from .editing_navigation import WordLink, WordBookmark
    from .editing_math import WordEquation
    from .editing_fields import WordList
    from .editing_stories import WordNote, WordSection
    from docx.blkcntnr import BlockItemContainer
    if object_id == "body":
        return editor.body
    for part in editor._backend.part.package.parts:
        root = getattr(part, "element", None)
        if root is None or etree.QName(root).localname not in STORIES:
            continue
        node = next((node for node in root.iter() if node.get(ID) == object_id), None)
        if node is None:
            continue
        name = etree.QName(node).localname
        owner = _owner(editor, node, part)
        if name == "p":
            return WordParagraph(owner, node)
        if name == "tbl":
            return WordTable(owner, node)
        if name == "drawing":
            return WordPicture(owner, node)
        if name == "hyperlink":
            return WordLink(owner, node)
        if name == "sdt":
            from .editing_content import WordContentControl
            return WordContentControl(owner, node)
        if name == "oMath":
            return WordEquation(owner, node)
        if name == "bookmarkStart":
            end = next((end for end in root.iter(_tag("bookmarkEnd")) if end.get(_tag("id")) == node.get(_tag("id"))), None)
            if end is not None:
                return WordBookmark(owner, node, end)
        if name in {"footnote", "endnote"}:
            return WordNote(editor, node, BlockItemContainer(node, part), kind=name)
        if name == "sectPr":
            return WordSection(editor, editor._backend.element.body, editor._backend, section=node)
        if name == "num":
            return WordList(editor, node)
        if name in {"fldSimple", "fldChar"}:
            field = next((field for field in owner.fields if field._node is node), None)
            if field is not None:
                from .editing_navigation import WordLinkField, link_instruction
                if link_instruction(field.instruction) is not None:
                    return WordLinkField(field)
                return field
        if name in {"body", "tc", "txbxContent", "hdr", "ftr"}:
            return WordContainer(editor, node, BlockItemContainer(node, part))
    raise ObjectNotFound(object_id)


METHODS = {
    "WordParagraph": {"set_text", "replace_text", "format", "format_text", "insert_hyperlink", "add_field", "remove_list", "add_note", "add_equation", "delete", "move"},
    "WordTable": {"set_cell_text", "format", "format_cell", "format_row", "format_layout", "add_row", "delete_row", "add_column", "delete_column", "merge", "split_cell", "split", "delete", "move"},
    "WordPicture": {"resize", "crop", "rotate", "position", "delete"},
    "WordLink": {"set_target", "set_text", "unlink"},
    "WordLinkField": {"set_target", "set_text", "unlink", "set_instruction", "set_result", "lock"},
    "WordBookmark": {"rename", "delete"},
    "WordField": {"set_instruction", "set_result", "lock", "replace_result_text", "rebuild_toc"},
    "WordContentControl": {"configure", "set_text", "unwrap", "delete", "move", "insert_paragraph", "append_paragraph", "insert_table", "replace_all", "insert_control", "wrap_control"},
    "WordEquation": {"set_script", "delete"},
    "WordList": {"restart"},
    "WordContainer": {"insert_paragraph", "append_paragraph", "insert_table", "replace_all", "wrap_control", "insert_control"},
    "WordSection": {"insert_paragraph", "append_paragraph", "insert_table", "replace_all", "format_page", "format_columns"},
    "WordNote": {"insert_paragraph", "append_paragraph", "insert_table", "replace_all", "format_reference", "delete"},
}
GLOBAL = {"create_style", "configure_style", "delete_style", "format_style", "format_theme", "create_list", "insert_section", "update_fields"}


def _id(handle):
    node = getattr(handle, "_section", None)
    if node is None:
        node = getattr(handle, "_node", None)
    if node is None:
        node = getattr(handle, "_root", None)
    return identify(node) if node is not None else None


def _records(editor):
    from .editing_navigation import link_fields
    fields = link_fields(editor._backend)
    records = []
    for part in editor._backend.part.package.parts:
        root = getattr(part, "element", None)
        if root is None or etree.QName(root).localname not in STORIES:
            continue
        for node in root.iter():
            identity = node.get(ID)
            if identity is None:
                continue
            parent = next((parent.get(ID) for parent in node.iterancestors() if parent.get(ID)), None)
            kind = etree.QName(node).localname
            value = {"id": identity, "kind": kind, "part": str(part.partname), "parent_id": parent,
                     "fingerprint": sha256(etree.tostring(node)).hexdigest()}
            if kind == "p":
                value["text"] = WordParagraph(_owner(editor, node, part), node).text
            elif kind == "tbl":
                table = WordTable(_owner(editor, node, part), node)._native
                value.update(rows=len(table.rows), columns=len(table.columns))
            elif kind == "hyperlink":
                handle = resolve_object(editor, identity)
                value.update(text=handle.text, target=handle.target)
            elif kind == "bookmarkStart":
                value["name"] = node.get(_tag("name"))
            elif kind == "sdt":
                value.update(tag=resolve_object(editor, identity).tag, text=resolve_object(editor, identity).text)
            if node in fields:
                target, text = fields[node]
                value.update(kind="hyperlink", storage_kind=kind, text=text, target=target)
            records.append(value)
    return records


class WordEndpoint:
    """Own one WordDocument; all writes require a matching SHA-256 revision.

    IDs are stored as ignorable OOXML attributes, so inserted/moved objects
    retain identity after DOCX save/reopen. Initialization enables IDs once.
    """

    def __init__(self, document):
        from .document import WordDocument
        if not isinstance(document, WordDocument):
            raise TypeError("WordEndpoint requires a WordDocument.")
        self.document = document
        self._lock = RLock()
        self._history = OrderedDict()
        self._revision = None
        self._state = None
        self.snapshot()

    @staticmethod
    def _prepare(editor):
        from .editing_navigation import containers
        list(containers(editor))  # Activate note parts so their IDs persist too.
        stamp(editor._backend)

    def _bytes(self):
        return self.document._source.to_bytes(self.document)

    @property
    def version(self):
        with self._lock:
            return self._synchronize()["version"]

    def _synchronize(self):
        revision = sha256(self._bytes()).hexdigest()
        if self._state is None or revision != self._revision:
            with self.document.edit_word_details() as editor:
                self._prepare(editor)
                objects = _records(editor)
            self._revision = sha256(self._bytes()).hexdigest()
            self._state = {"version": self._revision, "objects": objects}
            self._index = {item["id"]: item for item in objects}
            self._history[self._revision] = self._state
            self._history.move_to_end(self._revision)
            while len(self._history) > 16:
                self._history.popitem(last=False)
        return self._state

    def snapshot(self):
        with self._lock:
            return deepcopy(self._synchronize())

    def changes(self, since_version):
        """Return net changes from one of the last 16 observed revisions."""
        if not isinstance(since_version, str):
            raise ValueError("since_version must be a revision string.")
        with self._lock:
            state = self._synchronize()
            baseline = self._history.get(since_version)
            if baseline is None:
                return {"version": state["version"], "requires_snapshot": True}
            old = {item["id"]: item for item in baseline["objects"]}
            new = {item["id"]: item for item in state["objects"]}
            return {"version": state["version"], "requires_snapshot": False,
                    "upserted": deepcopy([item for key, item in new.items() if old.get(key) != item]),
                    "deleted": [key for key in old if key not in new]}

    def get_objects(self, object_ids):
        if not isinstance(object_ids, list) or len(object_ids) > 1000 or any(not isinstance(value, str) for value in object_ids):
            raise ValueError("Provide a list of at most 1000 object IDs.")
        with self._lock:
            state = self._synchronize()
            records = self._index
            ids = list(dict.fromkeys(object_ids))
            return {"version": state["version"], "objects": deepcopy([records[key] for key in ids if key in records]),
                    "missing": [key for key in ids if key not in records]}

    def get_object(self, object_id):
        snapshot = self.get_objects([object_id])
        if snapshot["missing"]:
            raise ObjectNotFound(object_id)
        return {"version": snapshot["version"], "object": snapshot["objects"][0]}

    def apply(self, expected_version, operations, *, compact=False):
        if not isinstance(compact, bool):
            raise ValueError("compact must be a boolean.")
        if not isinstance(expected_version, str) or not isinstance(operations, list) or len(operations) > 1000:
            raise ValueError("Provide expected_version and a list of at most 1000 operations.")
        if any(not isinstance(operation, dict) for operation in operations):
            raise ValueError("Each edit operation must be an object.")
        with self._lock:
            current = self.version
            if expected_version != current:
                raise VersionConflict("Document revision changed; synchronize changes before editing.")
            before = self.snapshot()
            aliases, results = {}, []
            with self.document.edit_word_details() as editor:
                self._prepare(editor)

                def resolve(value):
                    return editor.object(aliases.get(value, value))

                for operation in operations:
                    if operation.keys() - {"target", "method", "args", "kwargs", "as", "cell"}:
                        raise ValueError("Unknown edit operation keys.")
                    method = operation.get("method")
                    if not isinstance(operation.get("kwargs", {}), dict):
                        raise ValueError("Operation kwargs must be an object.")
                    args, kwargs = operation.get("args", []), dict(operation.get("kwargs", {}))
                    if not isinstance(args, list):
                        raise ValueError("Operation args must be a JSON array.")
                    target = operation.get("target", "document")
                    if target == "document":
                        if method not in GLOBAL:
                            raise ValueError("Unsupported document method.")
                        result = getattr(editor, method)(*args, **kwargs)
                    elif method in {"replace_range", "copy_range", "move_range", "bookmark_range", "hyperlink_range"}:
                        from .editing_ranges import WordRange
                        start, end = resolve(target), resolve(kwargs.pop("end_id", target))
                        if not isinstance(start, WordParagraph) or not isinstance(end, WordParagraph):
                            raise ValueError("Text range endpoints must be paragraphs.")
                        selection = WordRange(start, kwargs.pop("start", 0), end, kwargs.pop("end", len(end.text)))
                        if method == "replace_range":
                            result = selection.replace(**kwargs)
                        elif method == "bookmark_range":
                            result = selection.add_bookmark(**kwargs)
                        elif method == "hyperlink_range":
                            result = selection.add_hyperlink(**kwargs)
                        else:
                            destination = resolve(kwargs.pop("destination"))
                            result = selection.copy_to(destination, kwargs.pop("offset"), move=method == "move_range", **kwargs)
                    else:
                        handle = resolve(target)
                        if "cell" in operation:
                            if not isinstance(handle, WordTable):
                                raise ValueError("Only tables accept cell coordinates.")
                            handle = handle.cell(*operation["cell"])
                        if method in {"move", "move_reference"} and "destination" in kwargs:
                            kwargs["destination"] = resolve(kwargs["destination"])
                            if not isinstance(kwargs["destination"], WordContainer):
                                raise ValueError("Move destination must be a content container.")
                        if method == "set_list":
                            if not isinstance(handle, WordParagraph):
                                raise ValueError("Only paragraphs can use numbering.")
                            result = handle.set_list(resolve(kwargs.pop("listing")), **kwargs)
                        elif method == "join" and isinstance(handle, WordTable):
                            result = handle.join(resolve(kwargs.pop("following")), **kwargs)
                        elif method == "move_reference" and type(handle).__name__ == "WordNote":
                            result = handle.move_reference(resolve(kwargs.pop("paragraph")))
                        else:
                            if method not in METHODS.get(type(handle).__name__, set()):
                                raise ValueError("Unsupported object method.")
                            result = getattr(handle, method)(*args, **kwargs)
                    self._prepare(editor)
                    identity = _id(result)
                    if "as" in operation:
                        name = operation["as"]
                        if not isinstance(name, str) or not name.startswith("$") or name in aliases or identity is None:
                            raise ValueError("An operation alias must be a unique $name for a returned object.")
                        aliases[name] = identity
                    results.append({"id": identity} if identity is not None else {"value": result})
            after = self.snapshot()
            old = {value["id"]: value for value in before["objects"]}
            new = {value["id"]: value for value in after["objects"]}
            after.update(results=results, created=[key for key in new if key not in old], deleted=[key for key in old if key not in new],
                         changed=[key for key in new if key in old and new[key] != old[key]])
            if compact:
                after["upserted"] = [new[key] for key in [*after["created"], *after["changed"]]]
                del after["objects"]
            return after

    def export(self, format="docx"):
        with self._lock:
            if format == "docx":
                return self._bytes()
            if format == "hwpx":
                path = self.document._ensure_temp_dir() / "endpoint-export.hwpx"
                self.document.to_hwpx_document().save(path)
                return path.read_bytes()
            raise ValueError("Export format must be docx or hwpx.")


def make_http_server(endpoint, *, host="127.0.0.1", port=8765, token=None):
    """Return an unstarted HTTP server. Caller owns serve_forever/shutdown."""
    from hmac import compare_digest

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def _send(self, status, body, content_type="application/json"):
            if content_type == "application/json":
                body = json.dumps(body, ensure_ascii=False, allow_nan=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _handle(self, method):
            if token is not None and not compare_digest(self.headers.get("Authorization", ""), "Bearer " + token):
                self._send(401, {"error": "unauthorized"})
                return
            url = urlsplit(self.path)
            path = url.path
            try:
                if method == "GET" and path == "/document":
                    self._send(200, endpoint.snapshot())
                elif method == "GET" and path == "/version":
                    self._send(200, {"version": endpoint.version})
                elif method == "GET" and path == "/changes":
                    query = parse_qs(url.query, keep_blank_values=True)
                    if set(query) != {"since"} or len(query["since"]) != 1:
                        raise ValueError("Provide one since revision.")
                    self._send(200, endpoint.changes(query["since"][0]))
                elif method == "GET" and path.startswith("/objects/"):
                    self._send(200, endpoint.get_object(path.removeprefix("/objects/")))
                elif method == "GET" and path in {"/export.docx", "/export.hwpx"}:
                    self._send(200, endpoint.export(path.rsplit(".", 1)[1]), "application/octet-stream")
                elif method == "POST" and path in {"/edits", "/objects"}:
                    if self.headers.get("Content-Type", "").split(";", 1)[0].strip().lower() != "application/json":
                        raise ValueError("Edit requests require application/json.")
                    size = int(self.headers.get("Content-Length", "0"))
                    if not 0 < size <= 1024 * 1024:
                        raise ValueError("Request body must contain 1..1048576 bytes.")
                    def constant(value):
                        raise ValueError("Non-finite JSON values are not supported.")
                    data = json.loads(self.rfile.read(size), parse_constant=constant)
                    if path == "/objects":
                        if not isinstance(data, dict) or data.keys() != {"ids"}:
                            raise ValueError("Request needs ids.")
                        self._send(200, endpoint.get_objects(data["ids"]))
                        return
                    if not isinstance(data, dict) or not {"expected_version", "operations"} <= data.keys() or data.keys() - {"expected_version", "operations", "compact"}:
                        raise ValueError("Request needs expected_version and operations.")
                    self._send(200, endpoint.apply(data["expected_version"], data["operations"], compact=data.get("compact", False)))
                else:
                    self._send(404, {"error": "not_found"})
            except VersionConflict as error:
                self._send(409, {"error": str(error), "version": endpoint.version})
            except ObjectNotFound as error:
                self._send(404, {"error": "object_not_found", "id": str(error.args[0])})
            except (ValueError, TypeError, IndexError, KeyError, RuntimeError) as error:
                self._send(422, {"error": str(error)})
            except Exception:
                self._send(500, {"error": "internal_error"})

        def do_GET(self):
            self._handle("GET")

        def do_POST(self):
            self._handle("POST")

    return ThreadingHTTPServer((host, port), Handler)
