from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Callable, Literal

from .document import HwpxDocument
from .hancom_document import HancomDocument
from .hwp_document import HwpDocument

HancomConverter = Callable[[str | Path, str | Path, str], Path]
BridgeAuthority = Literal["hwp", "hwpx", "hancom"]


class HwpHwpxBridge:
    def __init__(
        self,
        *,
        hwp_document: HwpDocument | None = None,
        hwpx_document: HwpxDocument | None = None,
        hancom_document: HancomDocument | None = None,
        converter: HancomConverter | None = None,
    ) -> None:
        if hwp_document is None and hwpx_document is None and hancom_document is None:
            raise ValueError("At least one source document must be provided.")
        self._hwp_document = hwp_document
        self._hwpx_document = hwpx_document
        self._hancom_document = hancom_document
        self._converter = converter
        self._workspace: Path | None = None
        if hancom_document is not None:
            self._authority: BridgeAuthority = "hancom"
        elif hwpx_document is not None:
            self._authority = "hwpx"
        else:
            self._authority = "hwp"

    @classmethod
    def open(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HwpHwpxBridge":
        source_path = Path(path).expanduser().resolve()
        suffix = source_path.suffix.lower()
        if suffix == ".hwp":
            return cls.from_hwp(source_path, converter=converter)
        if suffix == ".hwpx":
            return cls.from_hwpx(source_path, converter=converter)
        raise ValueError(f"Unsupported bridge source: {source_path}")

    @classmethod
    def from_hwp(
        cls,
        source: str | Path | HwpDocument,
        *,
        converter: HancomConverter | None = None,
    ) -> "HwpHwpxBridge":
        if isinstance(source, HwpDocument):
            document = source
        else:
            document = HwpDocument.open(source, converter=converter)
        return cls(hwp_document=document, converter=converter)

    @classmethod
    def from_hwpx(
        cls,
        source: str | Path | HwpxDocument,
        *,
        converter: HancomConverter | None = None,
    ) -> "HwpHwpxBridge":
        if isinstance(source, HwpxDocument):
            document = source
        else:
            document = HwpxDocument.open(source)
        return cls(hwpx_document=document, converter=converter)

    @classmethod
    def from_hancom(
        cls,
        source: HancomDocument,
        *,
        converter: HancomConverter | None = None,
    ) -> "HwpHwpxBridge":
        return cls(hancom_document=source, converter=converter)

    def hancom_document(self, *, force_refresh: bool = False) -> HancomDocument:
        document = self._materialize_hancom_document(force_refresh=force_refresh)
        self._authority = "hancom"
        return document

    def hwp_document(self, *, force_refresh: bool = False) -> HwpDocument:
        document = self._materialize_hwp_document(force_refresh=force_refresh)
        self._authority = "hwp"
        return document

    def hwpx_document(self, *, force_refresh: bool = False) -> HwpxDocument:
        document = self._materialize_hwpx_document(force_refresh=force_refresh)
        self._authority = "hwpx"
        return document

    def refresh_hancom(self) -> HancomDocument:
        return self.hancom_document(force_refresh=True)

    def refresh_hwp(self) -> HwpDocument:
        return self.hwp_document(force_refresh=True)

    def refresh_hwpx(self) -> HwpxDocument:
        return self.hwpx_document(force_refresh=True)

    def save_hwp(self, path: str | Path, *, force_refresh: bool = False) -> Path:
        document = self._materialize_hwp_document(force_refresh=force_refresh or self._authority != "hwp")
        return document.save(path)

    def save_hwpx(self, path: str | Path, *, force_refresh: bool = False) -> Path:
        document = self._materialize_hwpx_document(force_refresh=force_refresh or self._authority != "hwpx")
        return document.save(path)

    def save(self, path: str | Path, *, force_refresh: bool = False) -> Path:
        target_path = Path(path).expanduser().resolve()
        suffix = target_path.suffix.lower()
        if suffix == ".hwp":
            return self.save_hwp(target_path, force_refresh=force_refresh)
        if suffix == ".hwpx":
            return self.save_hwpx(target_path, force_refresh=force_refresh)
        raise ValueError(f"Unsupported bridge target: {target_path}")

    def _materialize_hancom_document(self, *, force_refresh: bool = False) -> HancomDocument:
        if self._hancom_document is not None and not force_refresh:
            return self._hancom_document

        if self._authority == "hwpx" and self._hwpx_document is not None:
            self._hancom_document = self._hwpx_document.to_hancom_document(converter=self._converter)
            return self._hancom_document
        if self._authority == "hwp" and self._hwp_document is not None:
            self._hancom_document = self._hwp_document.to_hancom_document(force_refresh=force_refresh)
            return self._hancom_document
        if self._hwpx_document is not None:
            self._hancom_document = self._hwpx_document.to_hancom_document(converter=self._converter)
            return self._hancom_document
        if self._hwp_document is not None:
            self._hancom_document = self._hwp_document.to_hancom_document(force_refresh=force_refresh)
            return self._hancom_document
        raise ValueError("No source document is available to build a HancomDocument.")

    def _materialize_hwp_document(self, *, force_refresh: bool = False) -> HwpDocument:
        if self._hwp_document is not None and not force_refresh and self._authority == "hwp":
            return self._hwp_document
        if self._authority == "hancom" and self._hancom_document is not None:
            self._hwp_document = self._hancom_document.to_hwp_document(converter=self._converter)
            return self._hwp_document
        if self._authority == "hwpx" and self._hwpx_document is not None:
            self._hwp_document = self._hwpx_document.to_hancom_document(converter=self._converter).to_hwp_document(
                converter=self._converter
            )
            return self._hwp_document
        if self._hancom_document is not None:
            self._hwp_document = self._hancom_document.to_hwp_document(converter=self._converter)
            return self._hwp_document
        if self._hwpx_document is not None:
            self._hwp_document = self._hwpx_document.to_hancom_document(converter=self._converter).to_hwp_document(
                converter=self._converter
            )
            return self._hwp_document
        if self._hwp_document is not None:
            return self._hwp_document
        raise ValueError("No source document is available to build an HWP document.")

    def _materialize_hwpx_document(self, *, force_refresh: bool = False) -> HwpxDocument:
        if self._hwpx_document is not None and not force_refresh and self._authority == "hwpx":
            return self._hwpx_document
        if self._authority == "hancom" and self._hancom_document is not None:
            self._hwpx_document = self._hancom_document.to_hwpx_document()
            return self._hwpx_document
        if self._authority == "hwp" and self._hwp_document is not None:
            self._hwpx_document = self._hwp_document.to_hancom_document(force_refresh=force_refresh).to_hwpx_document()
            return self._hwpx_document
        if self._hancom_document is not None:
            self._hwpx_document = self._hancom_document.to_hwpx_document()
            return self._hwpx_document
        if self._hwp_document is not None:
            self._hwpx_document = self._hwp_document.to_hancom_document(force_refresh=force_refresh).to_hwpx_document()
            return self._hwpx_document
        if self._hwpx_document is not None:
            return self._hwpx_document
        raise ValueError("No source document is available to build an HWPX document.")

    def _ensure_workspace(self) -> Path:
        if self._workspace is None:
            self._workspace = Path(tempfile.mkdtemp(prefix="jakal_hwp_hwpx_bridge_"))
        return self._workspace
