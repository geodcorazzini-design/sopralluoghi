/* Compila il modello Word ASPI preparato con segnaposto {{...}}: testi, segni X rossi e foto. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.DocxFill = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const esc = s => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const multi = s => String(s ?? "").split(/\r?\n/).map(esc).join('</w:t><w:br/><w:t xml:space="preserve">');
  let nextId = 7000;

  function runAt(xml, idx) { const s = xml.lastIndexOf("<w:r>", idx); const e = xml.indexOf("</w:r>", idx) + 6; return [s, e]; }
  function paraAt(xml, idx) { const s = Math.max(xml.lastIndexOf("<w:p ", idx), xml.lastIndexOf("<w:p>", idx)); const e = xml.indexOf("</w:p>", idx) + 6; return [s, e]; }

  function lineRun(cx, cy, offh, offv, flip) {
    const id = nextId++;
    return `<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>` +
      `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251700000 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${offh}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${offv}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="19050" b="19050"/><wp:wrapNone/>` +
      `<wp:docPr id="${id}" name="Segno ${id}"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic xmlns:a="${A}"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvCnPr/>` +
      `<wps:spPr><a:xfrm${flip ? ' flipV="1"' : ""}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>` +
      `<a:ln w="22225"><a:solidFill><a:srgbClr val="EE0000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>` +
      `</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;
  }
  function photoRun(rid, cx, cy, name) {
    const id = nextId++;
    return `<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${esc(name)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic xmlns:a="${A}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }
  const capPara = t => `<w:p><w:pPr><w:spacing w:before="40" w:after="200"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Univers" w:hAnsi="Univers"/><w:i/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">${esc(t)}</w:t></w:r></w:p>`;

  /**
   * template: ArrayBuffer/Uint8Array del modello; fields: {TOKEN: testo}; marks: Set("AUT|D18","A.1|C",...);
   * photos: [{bytes: Uint8Array, w, h, caption}]; out: "blob" | "uint8array" | "nodebuffer"
   */
  async function fill(JSZip, template, fields, marks, photos, out) {
    const zip = await JSZip.loadAsync(template);
    let xml = await zip.file("word/document.xml").async("string");
    if (!xml.includes("{{DATA}}") || !xml.includes("{{FOTO}}")) throw new Error("Il file Word caricato non è il modello preparato per l'app (mancano i segnaposto).");
    nextId = 7000 + Math.floor(Math.random() * 1000);
    // segni X
    xml = xml.replace(/<w:r><w:t xml:space="preserve">\{\{X\|([^|]+)\|([^|]+)\|(\d+)\|(\d+)\|(-?\d+)\|(-?\d+)\}\}<\/w:t><\/w:r>/g,
      (m, grp, opt, cx, cy, offh, offv) => (marks.has(`${grp}|${opt}`) ? lineRun(cx, cy, offh, offv, true) + lineRun(cx, cy, offh, offv, false) : ""));
    // osservazioni e azioni: un paragrafo per riga (evita righe stirate dal giustificato)
    for (const k of ["OSSERVAZIONI", "AZIONI"]) {
      const i = xml.indexOf(`{{${k}}}`); if (i < 0) continue;
      const [s, e] = paraAt(xml, i); const para = xml.slice(s, e);
      const lines = String(fields[k] ?? "").split(/\r?\n/);
      xml = xml.slice(0, s) + lines.map(l => para.replace(`{{${k}}}`, esc(l))).join("") + xml.slice(e);
    }
    // altri testi (eventuali a capo con <w:br/>)
    xml = xml.replace(/\{\{([A-Z0-9_]+)\}\}/g, (m, k) => {
      if (k === "FOTO") return m;
      let v = fields[k];
      if ((k === "DATA" || k === "ORA_INIZIO" || k === "ORA_FINE") && !v) v = "_______________";
      return multi(v ?? "");
    });
    // foto: la sezione foto parte da una pagina nuova
    if (photos && photos.length) {
      const ti = xml.indexOf(fields.TITOLO_FOTO ? esc(fields.TITOLO_FOTO) : "{{TITOLO_FOTO}}");
      if (ti > 0) {
        let [tps] = paraAt(xml, ti);
        // togli i paragrafi vuoti prima del titolo: altrimenti possono creare una pagina bianca
        for (let k = 0; k < 6; k++) {
          const pe = tps; if (xml.slice(pe - 6, pe) !== "</w:p>") break;
          const ps = Math.max(xml.lastIndexOf("<w:p ", pe - 7), xml.lastIndexOf("<w:p>", pe - 7));
          const para = xml.slice(ps, pe);
          if (ps < 0 || /<w:t[ >]|<w:drawing|<w:br|<w:sectPr|<w:footnoteReference/.test(para)) break;
          xml = xml.slice(0, ps) + xml.slice(pe); tps = ps;
        }
        const pp = xml.indexOf("<w:pPr>", tps); const te = xml.indexOf("</w:p>", tps);
        if (pp > 0 && pp < te) xml = xml.slice(0, pp + 7) + "<w:pageBreakBefore/>" + xml.slice(pp + 7);
      }
    }
    const idx = xml.indexOf("{{FOTO}}");
    const [ps, pe] = paraAt(xml, idx);
    let body = "";
    if (photos && photos.length) {
      let rels = await zip.file("word/_rels/document.xml.rels").async("string");
      const MAXW = 2880000, MAXH = 3096000; // 8 cm x 8,6 cm: 4 foto per pagina (2 x 2)
      const runs = photos.map((p, i) => {
        const n = i + 1, rid = `rIdFotoApp${n}`, name = `foto_app_${n}.jpeg`;
        zip.file(`word/media/${name}`, p.bytes, { createFolders: false });
        rels = rels.replace("</Relationships>", `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/></Relationships>`);
        const k = Math.min(MAXW / p.w, MAXH / p.h);
        return { run: photoRun(rid, Math.round(p.w * k), Math.round(p.h * k), name), cap: p.caption || `Foto ${n}` };
      });
      for (let i = 0; i < runs.length; i += 2) {
        const pair = runs.slice(i, i + 2);
        const pb = i > 0 && i % 4 === 0 ? "<w:pageBreakBefore/>" : ""; // nuova pagina ogni 4 foto
        body += `<w:p><w:pPr><w:keepNext/>${pb}<w:spacing w:before="160" w:after="0"/><w:jc w:val="center"/></w:pPr>${pair.map(r => r.run).join('<w:r><w:t xml:space="preserve">    </w:t></w:r>')}</w:p>`;
        body += capPara(pair.map(r => r.cap).join("      ·      "));
      }
      zip.file("word/_rels/document.xml.rels", rels, { createFolders: false });
    } else body = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>`;
    xml = xml.slice(0, ps) + body + xml.slice(pe);
    zip.file("word/document.xml", xml, { createFolders: false });
    return zip.generateAsync({ type: out || "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", compression: "DEFLATE" });
  }
  return { fill };
});
