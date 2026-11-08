// MathML (DOM node) -> OMML string. Covers what MathLive emits for common math.
// ponytail: n-ary (∑ ∫) render as sub/sup-script, not m:nary; fences are plain runs, not m:d.
(function (root) {
  const esc = s => s.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const run = (t, plain) =>
    `<m:r>${plain ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''}<m:t xml:space="preserve">${esc(t)}</m:t></m:r>`;

  function conv(n) {
    if (!n) return '';
    const c = n.children, t = n.localName, x = i => conv(c[i]);
    const kids = () => [...c].map(conv).join('');
    switch (t) {
      case 'mi': return run(n.textContent, n.textContent.length > 1);
      case 'mn': return run(n.textContent);
      case 'mo': return run(n.textContent, /^[a-z]{2,}$/i.test(n.textContent)); // sin, lim: upright
      case 'mtext': return run(n.textContent, true);
      case 'mfrac': return `<m:f><m:num>${x(0)}</m:num><m:den>${x(1)}</m:den></m:f>`;
      case 'msup': return `<m:sSup><m:e>${x(0)}</m:e><m:sup>${x(1)}</m:sup></m:sSup>`;
      case 'msub': return `<m:sSub><m:e>${x(0)}</m:e><m:sub>${x(1)}</m:sub></m:sSub>`;
      case 'msubsup': case 'munderover':
        return `<m:sSubSup><m:e>${x(0)}</m:e><m:sub>${x(1)}</m:sub><m:sup>${x(2)}</m:sup></m:sSubSup>`;
      case 'munder': return `<m:limLow><m:e>${x(0)}</m:e><m:lim>${x(1)}</m:lim></m:limLow>`;
      case 'mover': return `<m:limUpp><m:e>${x(0)}</m:e><m:lim>${x(1)}</m:lim></m:limUpp>`;
      case 'msqrt': return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${kids()}</m:e></m:rad>`;
      case 'mroot': return `<m:rad><m:deg>${x(1)}</m:deg><m:e>${x(0)}</m:e></m:rad>`;
      default: return kids(); // math, mrow, mstyle, semantics...
    }
  }

  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
             'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
  const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  // Wraps one equation in the flat-OPC package that Word's insertOoxml expects.
  const toPackage = omml =>
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>` +
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>` +
    `</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
    `<w:document ${NS}><w:body><w:p><m:oMath>${omml}</m:oMath></w:p></w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`;

  const api = {
    mathmlToOmml: mathml => conv(new DOMParser().parseFromString(mathml, 'text/xml').documentElement),
    toPackage,
    conv,
  };
  if (typeof module !== 'undefined') module.exports = api; else root.Omml = api;
})(this);
