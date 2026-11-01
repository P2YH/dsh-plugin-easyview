/** Read Word body text from authorized original/checkpoint bytes, without mounting an editor. */
import JSZip from 'jszip';
import { DOMParser, type Node, type Element } from '@xmldom/xmldom';

// JSZip 3.10 exposes this API at runtime but omits it from JSZipObject's declarations.
declare module 'jszip' {
  interface JSZipObject { internalStream(type: 'uint8array'): JSZipStreamHelper<Uint8Array>; }
}

const WORD_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
]);
const MAX_XML_BYTES = 16 * 1024 * 1024;
const MAX_CONTEXT_CHARS = 256_000;

export async function readDocxContext(bytes: Uint8Array, signal: AbortSignal) {
  signal.throwIfAborted();
  const zip = await JSZip.loadAsync(bytes);
  signal.throwIfAborted();
  const body = zip.file('word/document.xml');
  if (!body) throw new Error('Word 文档缺少正文，无法读取。');
  // Bound decompression before building the XML tree, including highly compressed documents.
  const xml = await new Promise<string>((resolve, reject) => {
    const stream = body.internalStream('uint8array'), decoder = new TextDecoder('utf-8', { fatal: true });
    let content = '', size = 0, done = false;
    const fail = (error: unknown) => {
      if (done) return;
      done = true; stream.pause(); signal.removeEventListener('abort', abort); reject(error);
    };
    const abort = () => fail(signal.reason);
    signal.addEventListener('abort', abort, { once: true });
    stream.on('data', chunk => {
      if (done) return;
      size += chunk.byteLength;
      if (size > MAX_XML_BYTES) { fail(new Error('Word 正文过大，请拆分文档后再总结。')); return; }
      try { content += decoder.decode(chunk, { stream: true }); } catch (error) { fail(error); }
    }).on('error', fail).on('end', () => {
      if (done) return;
      try { content += decoder.decode(); } catch (error) { fail(error); return; }
      done = true; signal.removeEventListener('abort', abort); resolve(content);
    }).resume();
  });
  signal.throwIfAborted();
  const document = new DOMParser({ onError: () => { throw new Error('Word 正文 XML 无法解析。'); } }).parseFromString(xml, 'application/xml');
  if (document.doctype) throw new Error('Word 正文包含不支持的 XML 文档类型。');
  const root = document.documentElement;
  const namespace = root?.namespaceURI;
  if (!root || !namespace || !WORD_NAMESPACES.has(namespace) || root.localName !== 'document') throw new Error('Word 正文结构无效。');
  const bodyNode = root.getElementsByTagNameNS(namespace, 'body').item(0);
  if (!bodyNode) throw new Error('Word 文档缺少正文，无法读取。');
  const parts: string[] = [];
  // Iterative traversal also handles long/nested Word documents without recursive stack growth.
  const pending: Array<{ node: Node; closing?: boolean }> = [{ node: bodyNode }];
  while (pending.length) {
    const { node, closing } = pending.pop()!;
    if (node.nodeType !== 1) continue;
    const element = node as Element;
    const word = element.namespaceURI === namespace, tag = element.localName ?? '';
    if (closing) { parts.push(tag === 'tc' ? '\t' : '\n'); continue; }
    // Deleted revisions and field instructions are not visible body text.
    if (word && ['del', 'moveFrom', 'instrText', 'delText'].includes(tag)) continue;
    if (word && tag === 't') { parts.push(element.textContent ?? ''); continue; }
    if (word && ['tab', 'br', 'cr'].includes(tag)) { parts.push(tag === 'tab' ? '\t' : '\n'); continue; }
    if (word && ['p', 'tc', 'tr'].includes(tag)) pending.push({ node, closing: true });
    for (let child = node.lastChild; child; child = child.previousSibling) pending.push({ node: child });
  }
  const content = parts.join('').replace(/\n{3,}/g, '\n\n').trim();
  return { content: content.slice(0, MAX_CONTEXT_CHARS), truncated: content.length > MAX_CONTEXT_CHARS };
}
