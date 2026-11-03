import React, {useEffect, useState} from 'react';
import {createRoot} from 'react-dom/client';
import pptxgen from 'pptxgenjs';
import mammoth from 'mammoth';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import './style.css';
import './hh-brand-refresh.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

const C = {
  teal:'075766', berry:'A1164F', navy:'24205F', cyan:'2CB8CE',
  gold:'F5C84C', paper:'F7F5F0', pale:'EDF5F6', ink:'263238', white:'FFFFFF'
};
const layouts = [
  ['title','Title slide'],['agenda','Lesson agenda'],['section','Section divider'],
  ['objectives','Learning goals'],['content','Key points'],['summary','Quick summary'],
  ['activity','Class activity'],['discussion','Discussion'],['quiz','Knowledge check'],['chart','Data / chart slide']
];
const initial = [
  {type:'title',title:'Real skills. Real support. Real careers.',body:'A practical introduction to person-centred support',sub:'H&H Accredited Training  •  Faculty teaching deck'},
  {type:'agenda',title:'Today’s learning journey',body:'Understand the foundations\nExplore practical approaches\nApply learning to real situations\nCheck understanding'},
  {type:'objectives',title:'By the end of this lesson, you can…',body:'Explain the key principles\nIdentify safe and effective practice\nApply the approach to a case study'},
  {type:'content',title:'Start with the person',body:'Good support begins by understanding the person, their preferences and what matters to them.\n\nListen actively and communicate clearly.\nRespect choice, dignity and independence.\nAdapt support to the individual.'},
  {type:'activity',title:'Put it into practice',body:'In pairs, review the scenario and identify two ways to support choice and independence.\n\nAllow 5 minutes, then share one insight with the group.'},
  {type:'summary',title:'The takeaways',body:'Person-centred support starts with listening.\nSmall choices build confidence and independence.\nSafe practice and dignity go hand in hand.'}
];

function fileToDataUrl(file) {
  return new Promise((resolve,reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}
async function loadBrandLogo() {
  try {
    const response = await fetch('/assets/hh-logo.png');
    if (!response.ok) return null;
    return await fileToDataUrl(new File([await response.blob()], 'hh-logo.png', {type:'image/png'}));
  } catch { return null; }
}
const SAMPLE_PHOTOS = [
  ['hero','hero-support-team.jpg','Support in action'],
  ['childcare','childcare.jpg','Childcare'],
  ['agedCare','aged-care.jpg','Aged care & disability'],
  ['community','community-services.jpg','Community services'],
  ['business','business-leadership.jpg','Business & leadership'],
  ['firstAid','first-aid.jpg','First aid'],
  ['social','social-training.jpg','Social media'],
  ['courseGuide','course-guide.jpg','Course guide'],
  ['presentation','presentation-training.jpg','Presentation'],
  ['uniformTeal','uniform-teal.jpg','Teal uniform'],
  ['uniformNavy','uniform-navy.jpg','Navy uniform'],
  ['website','website-training.jpg','Website'],
  ['poster','poster-support.jpg','Poster']
];
const SAMPLE_PHOTO_DIMENSIONS = {
  hero:[328,377],childcare:[129,85],agedCare:[129,85],community:[122,85],business:[128,85],firstAid:[91,85],
  social:[123,231],courseGuide:[117,231],presentation:[177,180],uniformTeal:[132,129],uniformNavy:[135,129],
  website:[226,102],poster:[198,134]
};

async function loadSamplePhotos() {
  const photos = {};
  await Promise.all(SAMPLE_PHOTOS.map(async ([key,filename]) => {
    const response = await fetch(`/assets/hh-sample-photos/${filename}`);
    if (!response.ok) throw new Error(`A sample photograph could not be loaded: ${filename}`);
    const blob = await response.blob();
    photos[key] = await fileToDataUrl(new File([blob],filename,{type:'image/jpeg'}));
  }));
  return photos;
}

function brandBoardSlides(imageData, filename='H&H brand guidelines', photos={}) {
  const courseGallery = [
    ['childcare','Childcare'],['agedCare','Aged care & disability'],['community','Community services'],
    ['business','Business & leadership'],['firstAid','First aid']
  ].filter(([key])=>photos[key]).map(([key,label])=>({label,imageData:photos[key],imageWidth:SAMPLE_PHOTO_DIMENSIONS[key][0],imageHeight:SAMPLE_PHOTO_DIMENSIONS[key][1]}));
  const applicationGallery = [
    ['social','Social media'],['courseGuide','Course guide'],['presentation','Presentation slide'],
    ['uniformTeal','Teal uniform'],['uniformNavy','Navy uniform'],['website','Website header'],['poster','Poster']
  ].filter(([key])=>photos[key]).map(([key,label])=>({label,imageData:photos[key],imageWidth:SAMPLE_PHOTO_DIMENSIONS[key][0],imageHeight:SAMPLE_PHOTO_DIMENSIONS[key][1]}));
  return [
    {type:'title',title:'Real skills. Real support. Real careers.',body:'H&H Accredited Training',sub:'Brand identity and application guide',imageData:photos.hero},
    {type:'content',title:'A clear, confident brand promise',body:'Nationally recognised training in childcare, aged care, disability, community services and business.\n\nOur identity should feel clean, modern, friendly and highly readable. Every touchpoint should communicate practical skills, dedicated support and flexible study options.',imageData:photos.hero},
    {type:'content',title:'Logo suite and correct use',body:'Primary logo: horizontal H&H Accredited Training lock-up.\nStacked logo: use where vertical space is available.\nMonogram: compact H&H mark for small-format applications.\nRegulatory lock-up: include RTO 45704 where the context requires it.\n\nKeep the logo clear, legible and proportionate. Do not distort, recolour or crowd the mark.'},
    {type:'content',title:'Approved colour palette',body:'#075766  Deep teal — primary brand colour\n#A1164F  Berry — emphasis and calls to action\n#24205F  Navy — supporting depth\n#2CB8CE  Cyan — highlights and pathway accents\n#F5C84C  Gold — selective warmth and emphasis\n#F7F5F0  Warm white — primary light background\n#EDF5F6  Pale blue — panels and soft backgrounds\n#263238  Charcoal — readable body text'},
    {type:'content',title:'Typography: clear by design',body:'Manrope — headings, titles and prominent labels. Use ExtraBold, Bold and Semibold weights to establish hierarchy.\n\nSource Sans 3 — body copy, captions, supporting information and longer reading. Use Regular and Semibold weights.\n\nKeep line lengths comfortable, maintain strong contrast and avoid unnecessary typographic decoration.'},
    {type:'content',title:'Iconography and visual language',body:'Use simple, friendly, consistent line icons.\nPrefer open shapes, even stroke weights and recognisable forms.\nUse generous whitespace and restrained colour accents.\nPair photography with clear, human-centred messages.\nAvoid clutter, heavy effects and inconsistent illustration styles.'},
    {type:'gallery',title:'Our courses, brought to life',body:'Use authentic, people-centred photography to show the breadth of H&H training.',galleryItems:courseGallery},
    {type:'content',title:'The H&H learning pathway',body:'LEARN — Build real skills with expert trainers.\n\nCONNECT — Get support every step of the way.\n\nSUCCEED — Gain confidence and start your next career.\n\nUse this sequence to organise learner journeys, course communications and campaign messages.'},
    {type:'gallery',title:'One identity. Many touchpoints.',body:'The same visual identity carries across learner, digital and workplace experiences.',galleryItems:applicationGallery},
    {type:'content',title:'Brand applications and consistency',body:'Carry the same identity across social media, course guides, presentation slides, email signatures, certificates, uniforms, name badges, website headers and posters.\n\nKeep the logo, colour palette, typography, icon style and message hierarchy consistent. Adapt layouts to each format without changing the underlying identity.'},
    {type:'brandboard',title:'Original brand board — visual reference',body:`Source image: ${filename}\nThis slide preserves the supplied artwork as a visual reference. The preceding slides contain editable text, shapes and embedded photographs.`,imageData}
  ];
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
}

function App() {
  const [file,setFile] = useState(null);
  const [source,setSource] = useState('');
  const [slides,setSlides] = useState(initial);
  const [mode,setMode] = useState('Lecture');
  const [course,setCourse] = useState('Certificate III in Individual Support');
  const [busy,setBusy] = useState(false);
  const [active,setActive] = useState(0);
  const [notice,setNotice] = useState('Loading the H&H sample deck with embedded photography. Upload your own material to replace it.');

  const activeSlide = slides[active] || slides[0];
  const update = (key,value) => setSlides(items => items.map((slide,index) => index === active ? {...slide,[key]:value} : slide));
  const addSlide = type => {
    const body = type === 'quiz'
      ? 'Question: Add your question here\n\nA. Option one\nB. Option two\nC. Option three\nD. Option four'
      : type === 'discussion'
        ? 'What would you do in this situation?\n\nConsider the person’s needs, preferences and safety.'
        : 'Add your teaching points here.';
    setSlides(items => [...items,{type,title:layouts.find(item=>item[0]===type)?.[1] || 'New slide',body}]);
    setActive(slides.length);
  };

  function structure(text,kind) {
    const lines = text.replace(/\r/g,'').split('\n').map(line=>line.replace(/\u00a0/g,' ').replace(/[ \t]+/g,' ').trim());
    const chunks = [];
    let current = '';
    for (const line of lines) {
      if (!line) { if(current){chunks.push(current);current='';} continue; }
      if (/^(?:#{1,4}\s+|(?:\d+(?:\.\d+)*[.)]?\s+)|[A-Z][A-Z0-9 /&-]{4,})/.test(line) && line.length < 90) {
        if(current){chunks.push(current);current='';}
        chunks.push(line.replace(/^#{1,4}\s*/,''));
      } else {
        current += (current ? ' ' : '') + line;
        if (current.length > 650) { chunks.push(current); current=''; }
      }
    }
    if(current) chunks.push(current);
    const sections = [];
    let section = {title:'',items:[]};
    const isHeading = value => value.length < 76 && value.split(/\s+/).length <= 9 && !/[.!?]$/.test(value) && /^[A-Z0-9]/.test(value);
    for (const chunk of chunks) {
      if (isHeading(chunk) && section.items.length) {
        sections.push(section); section = {title:chunk.replace(/^[\d.]+[.)]?\s*/,''),items:[]};
      } else section.items.push(chunk);
    }
    if(section.items.length || section.title) sections.push(section);
    const classify = (title,body) => {
      const value = (title+' '+body).toLowerCase();
      if(/learning outcomes?|learning objectives?|by the end of|you will be able to/.test(value)) return 'objectives';
      if(/quiz|knowledge check|check your understanding|true or false|question\s*\d/.test(value) || (kind==='Quiz' && body.includes('?'))) return 'quiz';
      if(/discussion|reflect|think about/.test(value)) return 'discussion';
      if(/activity|practical task|group work|pair work|workshop/.test(value)) return 'activity';
      if(/summary|key takeaways|conclusion|recap/.test(value)) return 'summary';
      if(/agenda|session overview|lesson outline/.test(value)) return 'agenda';
      return 'content';
    };
    const title = (sections.find(item=>item.title)?.title || 'Teaching session').replace(/^[\d.)\s-]+/,'');
    const output = [{type:'title',title,body:course,sub:'H&H Accredited Training  •  Faculty teaching deck'}];
    const agenda = sections.find(item=>classify(item.title,item.items.join(' '))==='agenda');
    if (agenda) output.push({type:'agenda',title:agenda.title || 'Today’s learning journey',body:agenda.items.join('\n')});
    else {
      const headings = sections.filter(item=>item.title && classify(item.title,item.items.join(' '))==='content').map(item=>item.title);
      if(headings.length>=3 && kind!=='Quiz') output.push({type:'agenda',title:'Today’s learning journey',body:headings.slice(0,6).join('\n')});
    }
    sections.forEach(item=>{
      const body = item.items.join('\n');
      if(!body.trim()) return;
      const type = classify(item.title,body);
      if(type==='agenda') return;
      const groups = [];
      let group = [];
      let length = 0;
      item.items.forEach(part=>{
        if(group.length && (length+part.length>560 || group.length>=6)){groups.push(group);group=[];length=0;}
        group.push(part); length += part.length;
      });
      if(group.length) groups.push(group);
      groups.forEach((parts,index)=>{
        const fallback = {objectives:'Learning goals',quiz:'Knowledge check',discussion:'Discussion',activity:'Class activity',summary:'Quick summary',content:'Key teaching points'}[type];
        output.push({type,title:(item.title || fallback)+(groups.length>1?` (${index+1})`:''),body:parts.join('\n')});
      });
    });
    if(kind==='Quiz' && !output.some(item=>item.type==='quiz')) output.push({type:'quiz',title:'Knowledge check',body:'Add a question based on the material and include answer options for learners.'});
    if(kind==='Workshop' && !output.some(item=>['activity','discussion'].includes(item.type))) output.push({type:'activity',title:'Apply your learning',body:'Use the key ideas from this session to complete a practical task.'});
    if(!output.some(item=>item.type==='summary')) {
      const points=output.filter(item=>item.type==='content').slice(-3).map(item=>item.title);
      if(points.length) output.push({type:'summary',title:'Review and reflect',body:'Revisit these key areas:\n'+points.join('\n')+'\n\nWhat is one idea you will apply after this session?'});
    }
    return output;
  }

  async function loadBrandBoardSample() {
    setBusy(true); setNotice('Preparing the H&H brand-board sample…');
    try {
      const response=await fetch('/assets/hh-brand-guidelines.jpg');
      if(!response.ok) throw new Error('The bundled brand-board sample could not be loaded.');
      const blob=await response.blob();
      const data=await fileToDataUrl(new File([blob],'H&H-Accredited-Training-brand-board.jpg',{type:'image/jpeg'}));
      const photos=await loadSamplePhotos();
      const draft=brandBoardSlides(data,'H&H Accredited Training brand guidelines',photos);
      setFile({name:'H&H Accredited Training brand guidelines.png'});
      setSource('');
      setSlides(draft); setActive(0);
      setNotice(`Sample loaded: ${draft.length} slides. Edit the content, then export the PowerPoint to test the prototype.`);
    } catch(error) {
      setSlides([]); setActive(0);
      setNotice(`Sample could not be loaded: ${error.message || 'Please retry or upload your own material.'}`);
    }
    finally { setBusy(false); }
  }

  async function readFile(upload) {
    setBusy(true); setFile(upload); setNotice('Reading your material…');
    try {
      const ext=upload.name.split('.').pop().toLowerCase();
      let text='';
      if(['png','jpg','jpeg','webp'].includes(ext)) {
        const data=await fileToDataUrl(upload);
        const isHhBoard=/brand[\s_-]?(guidelines|board)/i.test(upload.name);
        const draft=isHhBoard
          ? brandBoardSlides(data,upload.name,await loadSamplePhotos())
          : [
              {type:'title',title:'Visual reference deck',body:'Use this presentation to discuss the supplied image and add your own teaching notes.',sub:upload.name},
              {type:'brandboard',title:'Original image — visual reference',body:`Source image: ${upload.name}\nThe image is preserved as a visual reference. Text in the image is not automatically extracted.`,imageData:data}
            ];
        setSource('');
        setSlides(draft); setActive(0);
        setNotice(isHhBoard
          ? `Imported ${upload.name}. Created ${draft.length} slides with editable brand summaries and the original image as a reference; review before export.`
          : `Imported ${upload.name} as a visual reference. Image text is not automatically extracted; add your own editable content before export.`);
        return;
      }
      if(['txt','md'].includes(ext)) text=await upload.text();
      else if(ext==='docx') text=(await mammoth.extractRawText({arrayBuffer:await upload.arrayBuffer()})).value;
      else if(ext==='pdf') {
        const pdf=await pdfjsLib.getDocument({data:await upload.arrayBuffer()}).promise;
        const pages=[];
        for(let pageNo=1;pageNo<=pdf.numPages;pageNo++) {
          const page=await pdf.getPage(pageNo);
          const content=await page.getTextContent();
          pages.push(content.items.map(item=>item.str).filter(Boolean).join(' '));
        }
        text=pages.join('\n\n');
      } else throw new Error('Please use PDF, DOCX, TXT, MD or an image (PNG, JPG, WEBP).');
      if(!text.trim()) throw new Error('No readable text was found. Scanned PDFs need OCR before upload.');
      setSource(text);
      const draft=structure(text,mode);
      setSlides(draft); setActive(0);
      setNotice(`Imported ${upload.name}. Created ${draft.length} draft slides. Review before export.`);
    } catch(error) { setNotice(error.message || 'Could not read that file.'); }
    finally { setBusy(false); }
  }

  async function exportPptx() {
    if(!slides.length) { setNotice('Load the H&H sample or upload source material before exporting.'); return; }
    setBusy(true); setNotice('Building your editable H&H PowerPoint…');
    try {
      const pptx=new pptxgen();
      const logo=await loadBrandLogo();
      pptx.layout='LAYOUT_WIDE';
      pptx.author='H&H Accredited Training';
      pptx.company='H&H Accredited Training';
      pptx.subject='H&H faculty teaching presentation';
      pptx.title=slides[0]?.title || 'H&H Teaching Slides';
      pptx.lang='en-AU';
      pptx.theme={headFontFace:'Manrope',bodyFontFace:'Source Sans 3',lang:'en-AU'};
      pptx.defineSlideMaster({
        title:'HH_BASE',
        background:{color:C.paper},
        objects:[
          {line:{x:0,y:7.32,w:13.333,h:0,line:{color:'D9E2E2',width:1}}},
          {text:{text:'H&H  |  ACCREDITED TRAINING',options:{x:.55,y:7.05,w:4,h:.18,fontFace:'Source Sans 3',fontSize:8,color:C.teal,bold:true,charSpacing:1}}},
          {text:{text:'REAL SKILLS. REAL SUPPORT. REAL CAREERS.',options:{x:8.25,y:7.05,w:4.5,h:.18,fontFace:'Source Sans 3',fontSize:7,color:'64757A',align:'right',charSpacing:.5}}}
        ],
        slideNumber:{x:12.8,y:7.05,color:C.teal,fontFace:'Source Sans 3',fontSize:8}
      });

      slides.forEach(item=>{
        const slide=pptx.addSlide('HH_BASE');
        if(item.type==='brandboard') {
          slide.background={color:C.paper};
          if(item.imageData) slide.addImage({data:item.imageData,x:1.05,y:.12,w:10.76,h:7.18});
          slide.addText('ORIGINAL BRAND BOARD  •  VISUAL REFERENCE',{x:.38,y:7.31,w:6,h:.1,fontFace:'Manrope',fontSize:6,bold:true,color:C.teal,margin:0,charSpacing:.6});
          return;
        }
        if(item.type==='gallery') {
          slide.background={color:C.paper};
          slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:.16,h:7,line:{color:C.teal,transparency:100},fill:{color:C.teal}});
          slide.addText(item.title,{x:.7,y:.48,w:11.8,h:.62,fontFace:'Manrope',fontSize:26,bold:true,color:C.teal,margin:0,fit:'shrink'});
          slide.addText(item.body||'',{x:.72,y:1.18,w:11.6,h:.35,fontFace:'Source Sans 3',fontSize:12,color:C.ink,margin:0,fit:'shrink'});
          const cards=item.galleryItems||[];
          const rows=cards.length>5?2:1;
          const perRow=rows===1?cards.length:4;
          const gap=.16;
          const cardW=rows===1?((11.85-gap*(cards.length-1))/cards.length):2.82;
          const cardH=rows===1?2.25:2.05;
          cards.forEach((card,index)=>{
            const row=rows===1?0:Math.floor(index/4);
            const col=rows===1?index:(row===0?index:index-4);
            const rowCount=rows===1?cards.length:(row===0?4:cards.length-4);
            const rowWidth=rowCount*cardW+(rowCount-1)*gap;
            const left=(13.333-rowWidth)/2;
            const x=left+col*(cardW+gap);
            const y=rows===1?2.35:1.88+row*2.35;
            slide.addShape(pptx.ShapeType.roundRect,{x,y,w:cardW,h:cardH+.42,rectRadius:.08,line:{color:'DCE7E7',width:1},fill:{color:C.white}});
            const maxImageW=cardW-.22, maxImageH=cardH-.55;
            const imageRatio=(card.imageWidth||1)/(card.imageHeight||1);
            const imageW=Math.min(maxImageW,maxImageH*imageRatio);
            const imageH=imageW/imageRatio;
            slide.addImage({data:card.imageData,x:x+(cardW-imageW)/2,y:y+.08+(maxImageH-imageH)/2,w:imageW,h:imageH});
            slide.addText(card.label,{x:x+.1,y:y+cardH-.38,w:cardW-.2,h:.3,fontFace:'Manrope',fontSize:rows===1?10:9,bold:true,color:C.teal,align:'center',valign:'mid',fit:'shrink',margin:0});
          });
          return;
        }
        const dark=['title','section'].includes(item.type);
        slide.background={color:dark?C.teal:C.paper};
        if(dark) {
          slide.addShape(pptx.ShapeType.roundRect,{x:.72,y:.42,w:2.48,h:1.08,rectRadius:.08,line:{color:C.white,transparency:100},fill:{color:C.white}});
          if(logo) slide.addImage({data:logo,x:.86,y:.51,w:2.2,h:.9});
          else {
            slide.addText('h&h',{x:.9,y:.57,w:1.7,h:.55,fontFace:'Manrope',fontSize:34,bold:true,color:C.teal,margin:0});
            slide.addText('Accredited Training',{x:.94,y:1.12,w:1.9,h:.18,fontFace:'Source Sans 3',fontSize:10,color:C.teal,margin:0});
          }
          if(item.imageData) {
            slide.addShape(pptx.ShapeType.roundRect,{x:8.35,y:1.55,w:4.25,h:4.95,rectRadius:.14,line:{color:C.cyan,width:1.2},fill:{color:C.teal}});
            slide.addImage({data:item.imageData,x:8.42,y:1.62,w:4.11,h:4.81});
          } else {
            slide.addShape(pptx.ShapeType.arc,{x:8.3,y:-.8,w:5.6,h:8.3,line:{color:C.cyan,width:3,transparency:5},rotate:5});
            slide.addShape(pptx.ShapeType.arc,{x:8.65,y:-.7,w:5.4,h:8.1,line:{color:C.berry,width:2,transparency:5},rotate:5});
          }
          slide.addText(item.title,{x:.8,y:2.35,w:7.2,h:1.65,fontFace:'Manrope',fontSize:30,bold:true,color:C.white,fit:'shrink',margin:0});
          slide.addText(item.body||'',{x:.82,y:4.18,w:6.2,h:1.2,fontFace:'Source Sans 3',fontSize:19,color:'EAF4F4',fit:'shrink',margin:0});
          slide.addText(item.sub||course,{x:.82,y:6.42,w:6.3,h:.3,fontFace:'Source Sans 3',fontSize:11,color:C.gold,margin:0});
        } else {
          slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:.16,h:7,line:{color:C.teal,transparency:100},fill:{color:C.teal}});
          if(logo) slide.addImage({data:logo,x:11.45,y:.22,w:1.25,h:.73});
          slide.addText(item.title,{x:.7,y:.55,w:10.2,h:.72,fontFace:'Manrope',fontSize:27,bold:true,color:C.teal,margin:0,fit:'shrink'});
          slide.addShape(pptx.ShapeType.line,{x:.72,y:1.42,w:1.2,h:0,line:{color:C.berry,width:3}});
          if(['agenda','objectives','summary'].includes(item.type)) {
            item.body.split(/\n+/).filter(Boolean).slice(0,6).forEach((text,index)=>{
              const y=1.9+index*.82;
              slide.addShape(pptx.ShapeType.ellipse,{x:.85,y:y+.1,w:.34,h:.34,line:{color:index%2?C.cyan:C.berry,transparency:100},fill:{color:index%2?C.cyan:C.berry}});
              slide.addText(text.replace(/^[-•\d.)\s]+/,''),{x:1.45,y,w:10.6,h:.55,fontFace:'Source Sans 3',fontSize:19,color:C.ink,fit:'shrink',margin:.03});
            });
          } else {
            slide.addShape(pptx.ShapeType.roundRect,{x:.75,y:1.85,w:11.8,h:4.75,rectRadius:.12,line:{color:'DCE7E7',width:1},fill:{color:C.white}});
            slide.addText(item.body,{x:1.15,y:2.2,w:item.imageData?6.1:10.9,h:4,fontFace:'Source Sans 3',fontSize:20,color:C.ink,fit:'shrink',breakLine:false,margin:.08});
            if(item.imageData) slide.addImage({data:item.imageData,x:7.65,y:2.2,w:4.45,h:3.8});
            if(item.resourceUrl) {
              slide.addShape(pptx.ShapeType.roundRect,{x:1.15,y:6,w:3.8,h:.42,rectRadius:.06,line:{color:C.teal,width:1},fill:{color:C.pale}});
              slide.addText('OPEN CLASS RESOURCE',{x:1.3,y:6.13,w:3.4,h:.12,fontFace:'Manrope',fontSize:8,bold:true,color:C.teal,hyperlink:{url:item.resourceUrl},margin:0});
            }
            if(['activity','discussion'].includes(item.type)) {
              slide.addShape(pptx.ShapeType.roundRect,{x:9.4,y:5.7,w:2.6,h:.5,rectRadius:.08,line:{color:C.teal,transparency:100},fill:{color:C.pale}});
              slide.addText(item.type==='activity'?'LEARN BY DOING':'SHARE YOUR THINKING',{x:9.52,y:5.86,w:2.35,h:.16,fontFace:'Manrope',fontSize:8,bold:true,color:C.teal,align:'center',margin:0});
            }
          }
        }
      });

      const presentationBlob=await pptx.write({outputType:'blob'});
      if(!presentationBlob || presentationBlob.size<12000) throw new Error('The PowerPoint export is incomplete. Please retry.');
      const signature=new Uint8Array(await presentationBlob.slice(0,4).arrayBuffer());
      if(signature[0]!==0x50 || signature[1]!==0x4B || signature[2]!==0x03 || signature[3]!==0x04) {
        throw new Error('The generated file is not a valid PowerPoint ZIP package. Please retry.');
      }
      const galleryImages=slides.filter(item=>item.type==='gallery').flatMap(item=>item.galleryItems||[]);
      if(galleryImages.some(item=>!item.imageData)) throw new Error('A gallery photograph is missing. Reload the H&H sample before exporting.');
      downloadBlob(presentationBlob,'HH-Accredited-Training-Sample.pptx');
      setNotice(`PowerPoint exported successfully: ${slides.length} slides, ${galleryImages.length} embedded gallery photographs. H&H styling is applied directly to editable slide elements.`);
    } catch(error) {
      console.error(error);
      setNotice('Export failed: '+(error.message || 'Please try again.'));
    } finally { setBusy(false); }
  }

  useEffect(() => { loadBrandBoardSample(); }, []);

  return <div className="app">
    <aside className="sidebar">
      <div className="brand"><img src="/assets/hh-logo.png" alt="H&H Accredited Training"/><span>SLIDE STUDIO</span></div>
      <div className="side-label">WORKSPACE</div>
      <button className="side-item selected">✦ <span>Presentation builder</span></button>
      <a className="side-item" href="/how-to.html" style={{textDecoration:'none'}}>⌁ <span>Staff quick guide</span></a>
      <button className="side-item" onClick={()=>setNotice('The H&H brand system is locked to the supplied guidelines.')}>◈ <span>Brand system</span><b>LOCKED</b></button>
      <div className="side-bottom"><div className="avatar">HH</div><div><strong>Faculty workspace</strong><small>H&H Accredited Training</small></div></div>
    </aside>
    <main className="main">
      <header className="topbar">
        <div><div className="eyebrow">FACULTY TOOLKIT / POWERPOINT GENERATOR</div><h1>Teaching, beautifully presented.</h1><p>Turn your teaching material into a consistent, classroom-ready deck.</p></div>
        <div className="top-actions"><span className="secure">● H&H BRAND APPLIED</span><button className="export" disabled={busy||!slides.length} onClick={exportPptx}>{busy?'Preparing…':'↓  Export PowerPoint'}</button></div>
      </header>
      <section className="workspace">
        <div className="left-panel">
          <div className="panel-head"><div><span className="step">01 / SOURCE MATERIAL</span><h2>Start with your content</h2><p>Drop in lecture notes, a course guide or a worksheet.</p></div><span className="file-icon">↥</span></div>
          <label className="dropzone" onDragOver={event=>event.preventDefault()} onDrop={event=>{event.preventDefault();if(event.dataTransfer.files[0])readFile(event.dataTransfer.files[0]);}}>
            <input type="file" accept=".pdf,.docx,.txt,.md,.png,.jpg,.jpeg,.webp" onChange={event=>event.target.files[0]&&readFile(event.target.files[0])}/>
            <span className="upload-icon">↑</span><strong>{file?file.name:'Drop your material or brand board here'}</strong><span>or <u>browse files</u> from your computer</span><small>PDF · DOCX · TXT · MD · PNG · JPG · WEBP</small>
          </label>
          <button className="sample-btn" onClick={loadBrandBoardSample} disabled={busy}>✦  Reload H&H brand-board sample (11 slides)</button>
          <div className="or-line"><span>DECK SETTINGS</span></div>
          <label className="field-label">Teaching format</label>
          <div className="mode-switch">{['Lecture','Quiz','Workshop'].map(value=><button key={value} className={mode===value?'active':''} onClick={()=>setMode(value)}>{value}</button>)}</div>
          <label className="field-label">Course / unit name</label><input className="text-input" value={course} onChange={event=>setCourse(event.target.value)}/>
          <div className="brand-note"><span>✦</span><div><strong>H&H design system is active</strong><small>Brand logo, approved H&H colours and fonts are applied directly to editable PowerPoint elements.</small></div></div>
          <button className="build-btn" disabled={!source||busy} onClick={()=>{setSlides(structure(source,mode));setActive(0);setNotice('Your content has been reorganised into a teaching deck. Review each slide before export.');}}>✦  Build slide outline</button>
        </div>
        <div className="right-panel">
          <div className="panel-head right-head"><div><span className="step">02 / DECK EDITOR</span><h2>Your teaching deck</h2><p>Review and refine the slide content before export.</p></div><span className="slide-count">{slides.length} SLIDES</span></div>
          <div className="editor">
            <div className="slide-list">{slides.map((item,index)=><button key={`${index}-${item.title}`} className={`slide-thumb ${active===index?'current':''}`} onClick={()=>setActive(index)}><span className="thumb-no">{String(index+1).padStart(2,'0')}</span><span className="thumb-title">{item.title}</span><span className="thumb-type">{layouts.find(layout=>layout[0]===item.type)?.[1]||'Content'}</span></button>)}<button className="add-slide" onClick={()=>addSlide('content')}>＋ Add slide</button></div>
            <div className="edit-area">{activeSlide&&<>
              <div className="preview-wrap">{activeSlide.type==='brandboard'&&activeSlide.imageData?<div className="mini-reference"><img src={activeSlide.imageData} alt="Original H&H brand board reference"/><span>ORIGINAL VISUAL REFERENCE</span></div>:activeSlide.type==='gallery'?<div className="mini-gallery-preview"><div className="mini-gallery-heading"><div className="mini-logo">h<span>&</span>h <small>Accredited Training</small></div><h3>{activeSlide.title}</h3><p>{activeSlide.body}</p></div><div className="mini-gallery-grid">{(activeSlide.galleryItems||[]).map((item,index)=><div className="mini-gallery-card" key={`${item.label}-${index}`}>{item.imageData?<img src={item.imageData} alt={item.label}/>:<div className="mini-gallery-placeholder">Photo unavailable</div>}<span>{item.label}</span></div>)}</div></div>:<div className={`mini-slide ${['title','section'].includes(activeSlide.type)?'dark':''}`}><div className="mini-logo">h<span>&</span>h <small>Accredited Training</small></div><div className="mini-rule"></div><h3>{activeSlide.title}</h3><p>{activeSlide.body.split('\n').filter(Boolean).slice(0,4).join('  •  ')}</p>{activeSlide.imageData&&<img className="mini-slide-photo" src={activeSlide.imageData} alt="Slide photograph"/>}<div className="mini-footer">REAL SKILLS. REAL SUPPORT. REAL CAREERS.</div></div>}</div>
              <div className="edit-fields">
                <div className="edit-head"><span>EDIT SELECTED SLIDE</span><select value={activeSlide.type} onChange={event=>update('type',event.target.value)}>{layouts.map(([value,label])=><option key={value} value={value}>{label}</option>)}</select></div>
                <label>Slide heading</label><input value={activeSlide.title} onChange={event=>update('title',event.target.value)}/>
                <label>Teaching content</label><textarea rows="7" value={activeSlide.body} onChange={event=>update('body',event.target.value)} placeholder="Add text, bullet points or questions…"/>
                <label>Slide image (optional)</label><input type="file" accept="image/png,image/jpeg" onChange={event=>{const image=event.target.files?.[0];if(image)fileToDataUrl(image).then(data=>update('imageData',data));}}/>
                {activeSlide.imageData&&<button className="delete-slide" onClick={()=>update('imageData',null)}>Remove image</button>}
                <label>Video / worksheet link (optional)</label><input value={activeSlide.resourceUrl||''} onChange={event=>update('resourceUrl',event.target.value)} placeholder="https://…"/>
                <div className="field-hint">Tip: Put each bullet or answer option on a new line. Add a photo to create a balanced text-and-image slide. The exported slide remains editable.</div>
                <button className="delete-slide" onClick={()=>{if(slides.length>1){setSlides(items=>items.filter((_,index)=>index!==active));setActive(Math.max(0,active-1));}}}>Remove slide</button>
              </div>
            </>}</div>
          </div>
        </div>
      </section>
      <div className="bottom-row">
        <div className="bottom-card"><span className="bottom-symbol">▧</span><div><strong>Designed for teaching</strong><small>Title, agenda, goals, content, activities, discussion, quiz and summary layouts.</small></div></div>
        <div className="bottom-card"><span className="bottom-symbol">◉</span><div><strong>One consistent identity</strong><small>H&H colours, Manrope headings, Source Sans 3 body text and branded slide furniture.</small></div></div>
        <div className="bottom-card"><span className="bottom-symbol">↗</span><div><strong>PowerPoint-ready</strong><small>Editable text and shapes, widescreen 16:9 format, reusable slide master and H&H colour styling.</small></div></div>
      </div>
      <footer className="status"><span className="status-dot"></span>{notice}<span className="status-right">H&H ACCREDITED TRAINING · SYDNEY</span></footer>
    </main>
  </div>;
}
createRoot(document.getElementById('root')).render(<App/>);
