/** Post-build public parser/writer contract checks; no customer documents. */
import assert from 'node:assert/strict';
import test from 'node:test';
import { readXlsx, writeXlsx, cloneSheet } from '../dist/index.mjs';
import { ZipReader } from '../dist/zip/reader.mjs';
import { ZipWriter } from '../dist/zip/writer.mjs';

const encode = new TextEncoder();
const decode = new TextDecoder();
const png = Uint8Array.from([137,80,78,71,13,10,26,10]);
const NS = 'http://schemas.openxmlformats.org';
const marker = (name, row, col, rowOff=0, colOff=0) =>
  `<xdr:${name}><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:${name}>`;
const extent = { cx: 1252822, cy: 962025 };
const pic = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Synthetic"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="10001" y="20003"/><a:ext cx="${extent.cx}" cy="${extent.cy}"/></a:xfrm></xdr:spPr></xdr:pic>`;
const from = marker('from', 0, 1, 257175, 133350);
const to = marker('to', 3, 8, 100965, 123158);
const two = (edit='') => `<xdr:twoCellAnchor${edit ? ` editAs="${edit}"` : ''}>${from}${to}${pic}<xdr:clientData/></xdr:twoCellAnchor>`;
const one = `<xdr:oneCellAnchor>${from}<xdr:ext cx="${extent.cx}" cy="${extent.cy}"/>${pic}<xdr:clientData/></xdr:oneCellAnchor>`;
const absolute = `<xdr:absoluteAnchor><xdr:pos x="-3175" y="19051"/><xdr:ext cx="${extent.cx}" cy="${extent.cy}"/>${pic}<xdr:clientData/></xdr:absoluteAnchor>`;

async function fixture(anchor, styleXml) {
  const original = await writeXlsx({ sheets: [{ name: 'Geometry', rows: [['ordinary cell']], images: [{data:png,type:'png',anchor:{from:{row:0,col:0}}}] }] });
  const zip = new ZipReader(original);
  const out = new ZipWriter();
  for (const path of zip.entries()) {
    let bytes = await zip.extract(path);
    if (path === 'xl/drawings/drawing1.xml') bytes = encode.encode(`<xdr:wsDr xmlns:xdr="${NS}/drawingml/2006/spreadsheetDrawing" xmlns:a="${NS}/drawingml/2006/main" xmlns:r="${NS}/officeDocument/2006/relationships">${anchor}</xdr:wsDr>`);
    if (path === 'xl/styles.xml' && styleXml) bytes = encode.encode(styleXml);
    if (path === 'xl/worksheets/sheet1.xml' && styleXml) bytes = encode.encode(decode.decode(bytes).replace('<c ', '<c s="0" '));
    out.add(path, bytes, {compress:false});
  }
  return out.build();
}
const imageOf = async bytes => (await readXlsx(bytes, {readStyles:true})).sheets[0].images[0];

for (const edit of ['', 'twoCell', 'oneCell', 'absolute']) {
  test(`read/write two-cell container retains markers and editAs=${edit || '(omitted)'}`, async () => {
    const image = await imageOf(await fixture(two(edit)));
    assert.equal(image.anchor.kind, 'twoCell');
    assert.equal(image.anchor.editAs, edit || undefined);
    assert.deepEqual(image.anchor.extent, extent);
    assert.deepEqual(image.anchor.from, {row:0,col:1,rowOff:257175,colOff:133350});
    assert.deepEqual(image.anchor.to, {row:3,col:8,rowOff:100965,colOff:123158});
    assert.equal(image.width, extent.cx / 9525);
    if (edit === 'absolute') assert.deepEqual(image.anchor.position, {x:10001,y:20003});
    const again = await imageOf(await writeXlsx({sheets:[{name:'Roundtrip',rows:[],images:[image]}]}));
    assert.deepEqual(again.anchor, image.anchor);
    assert.equal(again.width, image.width);
  });
}
for (const [kind, xml] of [['oneCell',one], ['absolute',absolute]]) {
  test(`${kind} container survives a read/write/read cycle without fabricated end marker`, async () => {
    const image = await imageOf(await fixture(xml));
    assert.equal(image.anchor.kind, kind);
    assert.equal(image.anchor.to, undefined);
    assert.deepEqual(image.anchor.extent, extent);
    if (kind === 'absolute') assert.deepEqual(image.anchor.position, {x:-3175,y:19051});
    const again = await imageOf(await writeXlsx({sheets:[{name:'Roundtrip',rows:[],images:[image]}]}));
    assert.deepEqual(again.anchor, image.anchor);
  });
}
test('one-cell extent is the container extent, not stale picture transform ext', async () => {
  const image = await imageOf(await fixture(one.replace(`<xdr:ext cx="${extent.cx}" cy="${extent.cy}"/>`, '<xdr:ext cx="1" cy="2"/>')));
  assert.deepEqual(image.anchor.extent,{cx:1,cy:2});
  assert.equal(image.width, 1 / 9525);
});
test('zero-size extents remain zero and do not turn into a default-size image', async () => {
  const input = await imageOf(await fixture(one.replace(`<xdr:ext cx="${extent.cx}" cy="${extent.cy}"/>`, '<xdr:ext cx="0" cy="0"/>')));
  assert.deepEqual(input.anchor.extent,{cx:0,cy:0});
  const output = await imageOf(await writeXlsx({sheets:[{name:'Zero',rows:[],images:[input]}]}));
  assert.deepEqual(output.anchor.extent,{cx:0,cy:0});
});
for (const value of ['NaN','Infinity','-1','9007199254740992','1.5']) {
  test(`rejects invalid saved EMU extent ${value}`, async () => {
    const image = await imageOf(await fixture(two().replace(`cx="${extent.cx}"`, `cx="${value}"`)));
    assert.equal(image.anchor.extent, undefined);
  });
}
test('explicit zero offsets and subpixel offsets survive serialization', async () => {
  const image = await imageOf(await fixture(two('oneCell')));
  image.anchor.from = {row:0,col:0,rowOff:0,colOff:1};
  image.anchor.to = {row:1,col:1,rowOff:2,colOff:0};
  const output = await imageOf(await writeXlsx({sheets:[{name:'Tiny',rows:[],images:[image]}]}));
  assert.deepEqual({...output.anchor.from,rowOff:output.anchor.from.rowOff ?? 0},image.anchor.from);
  assert.deepEqual({...output.anchor.to,colOff:output.anchor.to.colOff ?? 0},image.anchor.to);
});
test('cloneSheet isolates all nested drawing coordinates', async () => {
  const original = (await readXlsx(await fixture(two('absolute')))).sheets[0];
  const cloned = cloneSheet(original,'Copy');
  cloned.images[0].anchor.extent.cx++;
  cloned.images[0].anchor.position.x++;
  cloned.images[0].anchor.from.rowOff++;
  assert.equal(original.images[0].anchor.extent.cx,extent.cx);
  assert.equal(original.images[0].anchor.position.x,10001);
  assert.equal(original.images[0].anchor.from.rowOff,257175);
});
test('structuredClone preserves the full worker-safe geometry contract', async () => {
  const image = await imageOf(await fixture(two('absolute')));
  assert.deepEqual(structuredClone(image).anchor,image.anchor);
});
function styles(xfId='1', fontId='2') {
  return `<styleSheet xmlns="${NS}/spreadsheetml/2006/main"><fonts count="3"><font><name val="Fallback"/><sz val="11"/></font><font><name val="Cell"/><sz val="8"/></font><font><name val="Normal Font"/><sz val="10"/><b/><color rgb="FF112233"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="2"><xf fontId="0"/><xf fontId="${fontId}"/></cellStyleXfs><cellXfs count="1"><xf fontId="1" fillId="0" borderId="0" numFmtId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Localized name" builtinId="0" xfId="${xfId}"/></cellStyles></styleSheet>`;
}
test('Normal font follows builtinId -> xfId -> fontId, not fonts[0] or the first cell', async () => {
  const book = await readXlsx(await fixture(two(),styles()), {readStyles:true});
  assert.equal(book.defaultFont.name,'Normal Font');
  assert.equal(book.defaultFont.size,10);
  assert.equal(book.defaultFont.bold,true);
  assert.equal(book.sheets[0].cells.get('0,0').style.font.name,'Cell');
});
for (const [xf,font] of [['999','2'],['-1','2'],['bad','2'],['1','999'],['1','bad']]) {
  test(`invalid Normal reference ${xf}/${font} falls back safely`, async () => {
    const book = await readXlsx(await fixture(two(),styles(xf,font)), {readStyles:true});
    assert.equal(book.defaultFont.name,'Fallback');
  });
}
test('legacy caller-created image anchors remain supported', async () => {
  const image = {data:png,type:'png',anchor:{from:{row:1,col:2},to:{row:4,col:5}},width:17,height:23};
  const bytes = await writeXlsx({sheets:[{name:'Legacy',rows:[],images:[image]}]});
  const saved = await imageOf(bytes);
  assert.equal(saved.anchor.kind,'twoCell');
  assert.deepEqual(saved.anchor.extent,{cx:17*9525,cy:23*9525});
  const xml = decode.decode(await new ZipReader(bytes).extract('xl/drawings/drawing1.xml'));
  assert.ok(!xml.includes('NaN') && !xml.includes('undefined'));
});
