/** ExcelJS's writeBuffer() returns a real ArrayBuffer in the browser build but a Node
 * Buffer (a Uint8Array subclass) when running under Node (our unit tests) — this
 * normalizes either into a plain ArrayBuffer. */
export function toArrayBuffer(buf: ArrayBuffer | ArrayBufferView): ArrayBuffer {
  if (buf instanceof ArrayBuffer) return buf;
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}
