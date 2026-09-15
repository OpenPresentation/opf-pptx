import assert from 'node:assert/strict';
import {unzipSync} from 'fflate';
import {XMLParser} from 'fast-xml-parser';
import {toPptx,fromPptx} from '../dist/index.js';
import {prepareNodeFonts} from '@openpresentation/opf-render/fonts-node';
import {loadHarfBuzzShaper} from '@openpresentation/opf-render/font-shaping';

const parser=new XMLParser({ignoreAttributes:false,trimValues:false,parseTagValue:false});
const array=value=>value===undefined?[]:Array.isArray(value)?value:[value];
const runs=[{text:'of',fontSize:24,color:'#CC2222',underline:true,link:'https://example.org'},{text:'fice',fontSize:24,color:'#2222CC',strikethrough:true}];
const document={design:{fontScheme:{id:'roboto',heading:{family:'Gelasio'},body:{family:'Gelasio'},code:{family:'Gelasio'}}},slides:[{text:runs,notes:'Preserve notes'},{items:[runs]},{table:{columns:['Content'],rows:[[runs]]}}]};
const original=structuredClone(document),prepared=await prepareNodeFonts({pack:'office',fontShaper:await loadHarfBuzzShaper()});
try{
 const bytes=await toPptx(document,prepared.options),entries=unzipSync(bytes);
 for(let index=1;index<=3;index++){
  const xml=new TextDecoder().decode(entries[`ppt/slides/slide${index}.xml`]);
  const paragraphs=[];
  const visit=value=>{if(!value||typeof value!=='object')return;for(const [key,child]of Object.entries(value)){if(key==='a:p')paragraphs.push(...array(child));visit(child);}};
  visit(parser.parse(xml));
  const paragraph=paragraphs.find(p=>array(p['a:r']).map(run=>run['a:t']??'').join('')==='office');
  assert.ok(paragraph,`Slide ${index}: native editable text retains all source text`);
  const actual=array(paragraph['a:r']);assert.deepEqual(actual.map(run=>run['a:t']),['of','fice'],'Export retains the authored formatting runs within one paragraph');
  assert.deepEqual(actual.map(run=>run['a:rPr']['@_sz']),['2400','2400']);
  assert.deepEqual(actual.map(run=>run['a:rPr']['a:solidFill']['a:srgbClr']['@_val']),['CC2222','2222CC']);
  assert.equal(actual[0]['a:rPr']['@_u'],'sng');assert.ok(actual[0]['a:rPr']['a:hlinkClick']);
  assert.equal(actual[1]['a:rPr']['@_strike'],'sngStrike');
 }
 const imported=await fromPptx(bytes);assert.ok(JSON.stringify(imported).includes('office'));
 assert.deepEqual(document,original);
 console.log('Joint shaping exports original native body/list/table runs, text, sizes, colors, underline, strike and hyperlink without mutating source. Native Office layout remains a separate gate.');
}finally{prepared.registry.dispose();}
