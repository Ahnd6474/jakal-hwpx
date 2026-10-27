// Reset, measure, and apply in separate passes so a long document needs one
// layout calculation instead of a write/read cycle for every paragraph.
export function fitTextareas(elements: Iterable<HTMLTextAreaElement>): void {
  const fields = [...elements];
  for (const field of fields) field.style.height = '0px';
  const heights = fields.map(field => field.scrollHeight);
  fields.forEach((field, index) => { field.style.height = `${heights[index]}px`; });
}
