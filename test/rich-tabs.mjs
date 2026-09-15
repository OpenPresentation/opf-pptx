import assert from 'node:assert/strict';
import {unzipSync} from 'fflate';
import {XMLParser} from 'fast-xml-parser';
import {toPptx,fromPptx} from '../dist/index.js';
import {prepareNodeFonts} from '@openpresentation/opf-render/fonts-node';
import {loadHarfBuzzShaper} from '@openpresentation/opf-render/font-shaping';
import {resolvePresentation} from '@openpresentation/opf-render/svg';

const parser=new XMLParser({ignoreAttributes:false,trimValues:false,parseTagValue:false});
const array=value=>value===undefined?[]:Array.isArray(value)?value:[value];
const body=[{text:' A\t',fontSize:24,color:'#123456',link:'https://example.org'},{text:'\tB  \r\n\tC',fontSize:18,underline:true}];
const document={name:'Keep metadata',design:{fontScheme:{id:'roboto',heading:{family:'Arimo'},body:{family:'Arimo'},code:{family:'Arimo'}}},slides:[{text:body,notes:'Keep notes'},{items:[body]}]};
const original=structuredClone(document);let cases=0;
for(const mode of ['estimated','measured','painted']){
 const prepared=mode==='estimated'?null:await prepareNodeFonts({pack:'office',...(mode==='painted'?{fontShaper:await loadHarfBuzzShaper()}:{})});
 try{
  const options=prepared?.options??{},geometry=resolvePresentation(document,options);
  const bytes=await toPptx(document,options),entries=unzipSync(bytes);
  for(const [slideIndex,slide]of geometry.slides.entries()){
   const xml=new TextDecoder().decode(entries[`ppt/slides/slide${slideIndex+1}.xml`]);
   const shapes=array(parser.parse(xml)['p:sld']['p:cSld']['p:spTree']['p:sp']);
   const paragraphs=shapes.flatMap(shape=>array(shape['p:txBody']?.['a:p']));
   const lines=slide.geometry.items.flatMap(item=>item.text?.richLines??item.text?.listEntries?.flatMap(entry=>entry.text.richLines)??[]);
   assert.ok(lines.length,mode+' body/list lines exist');
   for(const line of lines){
    const source=line.fragments.map(fragment=>fragment.text).join('');
    const paragraph=paragraphs.find(p=>array(p['a:r']).map(run=>run['a:t']??'').join('')===source);
    assert.ok(paragraph,mode+' preserves native text including tabs and surrounding spaces');
    const stops=array(paragraph['a:pPr']?.['a:tabLst']?.['a:tab']).map(tab=>Number(tab['@_pos']));
    const expected=line.fragments.filter(fragment=>fragment.kind==='tab').map(fragment=>Math.round((fragment.x+fragment.width)*9525));
    assert.deepEqual(stops,expected,mode+' native paragraph uses the accepted line-relative tab stops');cases++;
   }
  }
  const imported=await fromPptx(bytes);
  assert.ok(JSON.stringify(imported).includes('\\t'),'Reimport retains current native tab characters');
  assert.deepEqual(document,original,'Export never rewrites authored runs or metadata');
 }finally{prepared?.registry.dispose();}
}
console.log(`Rich tabs: ${cases} estimated/measured/painted body/list paragraphs preserve source characters and accepted DrawingML tab stops. Native Office placement remains a separate gate.`);
