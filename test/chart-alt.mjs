// FA-09: Chart.alt is the chart frame's alternative text (p:nvGraphicFramePr/p:cNvPr/@descr) on classic and chartex (Office 2016)
// charts, "" is PowerPoint's decorative marker (adec:decorative), and fromPptx reads both back into chart.alt.
import assert from 'node:assert/strict';
import {unzipSync} from 'fflate';
import {XMLParser} from 'fast-xml-parser';
import {toPptx, fromPptx} from '../dist/index.js';
import {readFrameAlt, writeFrameAlt} from '../src/chart-alt.js';

const decoder = new TextDecoder();
const data = {columns: ['Quarter', 'North', 'South'], rows: [['Q1', 10, 5], ['Q2', 20, 8], ['Q3', 15, 12]]};
const parser = new XMLParser({ignoreAttributes: false, attributeNamePrefix: '', preserveOrder: false});

async function roundTrip(chart, slide = {}) {
  const bytes = await toPptx({design: {fontScheme: 'roboto'}, slides: [{title: 'Chart', ...slide, chart}]});
  const entries = unzipSync(bytes);
  const slideXml = decoder.decode(entries['ppt/slides/slide1.xml']);
  const imported = (await fromPptx(bytes)).slides[0];
  const back = imported.chart ?? imported.blocks?.find(block => block.chart)?.chart;
  return {slideXml, back, frames: [...slideXml.matchAll(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g)].map(match => match[0])};
}

const alt = 'Revenue grew from 10 in Q1 to 20 in Q2 & fell to 15 in Q3 ("North" leads).';
const escaped = 'Revenue grew from 10 in Q1 to 20 in Q2 &amp; fell to 15 in Q3 (&quot;North&quot; leads).';

// 1. Classic charts: descr on the frame, read back.
for (const type of ['column', 'bar', 'line', 'pie', 'scatter']) {
  const {slideXml, back, frames} = await roundTrip({type, alt, data});
  assert.equal(frames.length, 1, `${type}: one chart frame`);
  assert.match(frames[0], new RegExp(`<p:nvGraphicFramePr><p:cNvPr id="\\d+" name="OPF chart 1" descr="${escaped.replace(/[()]/g, '\\$&')}"/>`), `${type}: descr on the chart frame`);
  assert.doesNotThrow(() => parser.parse(slideXml), `${type}: well-formed slide`);
  assert.equal(back.alt, alt, `${type}: import restores alt`);
}

// 2. Chartex (Office 2016) charts: both the choice and the fallback frame carry the description.
for (const type of ['waterfall', 'funnel', 'treemap']) {
  const {slideXml, back, frames} = await roundTrip({type, alt, data: {columns: ['Step', 'Value'], rows: [['A', 10], ['B', 20], ['C', 15]]}});
  assert.match(slideXml, /<mc:AlternateContent/, `${type}: a chartex frame`);
  assert.equal(frames.length, 2, `${type}: choice and fallback frames`);
  for (const frame of frames) assert.match(frame, new RegExp(`descr="${escaped.replace(/[()]/g, '\\$&')}"`), `${type}: descr on the frame`);
  assert.equal(back.alt, alt, `${type}: import restores alt from the chartex choice`);
}

// 3. No alt: no descr and no decorative marker, and import invents none.
{
  const {slideXml, back} = await roundTrip({type: 'column', data});
  assert.doesNotMatch(slideXml, /<p:cNvPr[^>]*name="OPF chart 1"[^>]*descr=/, 'no alt, no descr');
  assert.doesNotMatch(slideXml, /adec:decorative/);
  assert.equal('alt' in back, false);
}

// 4. alt "" is PowerPoint's decorative marker, not an empty descr (PowerPoint writes descr="" for any shape without alt text).
for (const type of ['column', 'waterfall']) {
  const {slideXml, back, frames} = await roundTrip({type, alt: '', data: type === 'waterfall' ? {columns: ['Step', 'Value'], rows: [['A', 10], ['B', 20]]} : data});
  assert.doesNotMatch(slideXml, /descr=/, `${type}: no descr for a decorative chart`);
  for (const frame of frames) assert.match(frame, /<p:cNvPr id="\d+" name="OPF chart 1"><a:extLst><a:ext uri="\{C183D7F6-B498-43B3-948B-1728B52AA6E4\}"><adec:decorative xmlns:adec="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2017\/decorative" val="1"\/><\/a:ext><\/a:extLst><\/p:cNvPr>/);
  assert.doesNotThrow(() => parser.parse(slideXml), `${type}: well-formed slide`);
  assert.equal(back.alt, '', `${type}: import restores the empty alt`);
}

// 5. A PowerPoint chart frame that has an empty descr (no alt text) imports with no alt.
{
  const frame = '<p:nvGraphicFramePr><p:cNvPr id="4" name="Chart 3" descr=""/></p:nvGraphicFramePr>';
  assert.equal(readFrameAlt(parser.parse(frame)['p:nvGraphicFramePr']['p:cNvPr']), undefined);
  assert.equal(readFrameAlt(parser.parse('<p:cNvPr id="4" name="c" descr="Sales by region"/>')['p:cNvPr']), 'Sales by region');
}

// 6. writeFrameAlt keeps an existing cNvPr body, replaces an earlier descr and removes a stale decorative marker.
{
  const frame = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="c" descr="old"><a:hlinkClick r:id="rId2"/></p:cNvPr></p:nvGraphicFramePr></p:graphicFrame>';
  assert.match(writeFrameAlt(frame, 'new'), /<p:cNvPr id="4" name="c" descr="new"><a:hlinkClick r:id="rId2"\/><\/p:cNvPr>/);
  const marked = writeFrameAlt(frame, '');
  assert.match(marked, /<a:hlinkClick r:id="rId2"\/><a:extLst>/);
  assert.doesNotMatch(marked, /descr=/);
  const unmarked = writeFrameAlt(marked, 'now text');
  assert.doesNotMatch(unmarked, /adec:decorative/);
  assert.match(unmarked, /descr="now text"/);
}

console.log('chart alt: classic and chartex frames, decorative marker, import');
