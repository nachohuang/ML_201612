/** A minimal structural subset of the DOM API (just what word-adapter.ts needs), so the
 * same code type-checks against both the real browser DOMParser/XMLSerializer globals
 * (which satisfy this trivially — lib.dom.d.ts's Document/Element are a superset) and
 * @xmldom/xmldom's implementation used in Node tests (whose own Document/Element types
 * don't structurally match lib.dom.d.ts closely enough for TypeScript to accept one in
 * place of the other, even though both work identically at runtime for this subset).
 */

export interface MinimalNode {
  nodeType: number;
  parentNode: MinimalNode | null;
  nextSibling: MinimalNode | null;
  childNodes: ArrayLike<MinimalNode>;
  textContent: string | null;
  appendChild(node: MinimalNode): MinimalNode;
  insertBefore(node: MinimalNode, ref: MinimalNode | null): MinimalNode;
  removeChild(node: MinimalNode): MinimalNode;
}

export interface MinimalElement extends MinimalNode {
  localName: string | null;
  setAttribute(name: string, value: string): void;
  getElementsByTagNameNS(ns: string, localName: string): ArrayLike<MinimalElement>;
  getElementsByTagName(name: string): ArrayLike<MinimalElement>;
}

export interface MinimalDocument extends MinimalNode {
  getElementsByTagNameNS(ns: string, localName: string): ArrayLike<MinimalElement>;
  getElementsByTagName(name: string): ArrayLike<MinimalElement>;
  createElementNS(ns: string, qualifiedName: string): MinimalElement;
}

export interface MinimalDOMParser {
  parseFromString(text: string, mimeType: string): MinimalDocument;
}

export interface MinimalXMLSerializer {
  serializeToString(node: MinimalNode): string;
}
