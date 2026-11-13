/** The public template contains no employee data. All overlays are drawn here. */
function masaluSyncPersonalMenus_(page, assets, token, properties, requestedUid) {
  const data = masaluLoad_();
  const calendar = masaluBootstrap_(data);
  const shifts = new Map(calendar.shifts.filter(shift => !shift.isOff).map(shift => [shift.id, shift.name]));
  const counts = new Map();
  const employees = data.employees.rows.map(row => ({id:String(row[0] || '').trim(), uid:String(row[3] || '').trim()}));
  for (const person of employees) {
    if (person.id && validLineUid_(person.uid)) counts.set(person.uid, (counts.get(person.uid) || 0) + 1);
  }
  const targets = new Map(employees.filter(person => person.id && counts.get(person.uid) === 1)
    .map(person => [person.uid, person.id]));
  const assignments = new Map();
  for (const row of calendar.schedules) {
    if (row.date < page.start || row.date > page.end || !shifts.has(row.shiftId)) continue;
    if (!assignments.has(row.staffId)) assignments.set(row.staffId, []);
    const name = shifts.get(row.shiftId);
    const site = name.startsWith('玉門') ? '玉門' : name.startsWith('田園') ? '田園' : '其他';
    assignments.get(row.staffId).push([row.date, site, row.replyStatus === '同意']);
  }
  const prefix = 'MASALU_PERSONAL_MENU_';
  const ownRows = requestedUid ? assignments.get(targets.get(requestedUid)) || [] : [];
  const ownEntries = requestedUid ? masaluPersonalMenuEntries_(page,ownRows) : [];
  const saved = properties.getProperties();
  let template;
  let failures = 0;
  const users = requestedUid ? [requestedUid] :
    new Set([...targets.keys(), ...Object.keys(saved).filter(key => key.startsWith(prefix)).map(key => key.slice(prefix.length))]);
  for (const uid of users) {
    if (!validLineUid_(uid)) continue;
    const key = prefix + uid;
    const rows = (assignments.get(targets.get(uid)) || []).sort();
    try {
      const state = JSON.parse(saved[key] || '{}');
      if (state.retired) {
        masaluCalendarRichMenuRequest_('richmenu/'+state.retired,token,'delete',null,404);
        delete state.retired;
        properties.setProperty(key, JSON.stringify(state));
      }
      const signature = rows.length ? Utilities.base64EncodeWebSafe(
        Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify([page.start, page.personal, rows]))
      ) : '';
      const userPath = 'user/'+uid+'/richmenu';
      if (signature && state.signature === signature && state.id && !state.pending) {
        if (!requestedUid) continue;
        if (masaluCalendarMenuLinked_(userPath,state.id,token)) return {status:'linked',entries:ownEntries};
        // The account or menu was changed outside this script. Rebuild it.
      }
      if (!signature) {
        if (!state.id && !state.pending) continue;
        const linked = masaluCalendarRichMenuRequest_(userPath,token,'get',null,404);
        if (linked.getResponseCode() === 200 &&
            [state.id,state.pending?.id].includes(JSON.parse(linked.getContentText()).richMenuId)) {
          masaluCalendarRichMenuRequest_(userPath,token,'delete',null,404);
        }
        for (const id of new Set([state.id,state.pending?.id].filter(Boolean))) {
          masaluCalendarRichMenuRequest_('richmenu/'+id,token,'delete',null,404);
        }
        properties.deleteProperty(key);
        continue;
      }
      if (state.pending && state.pending.signature !== signature) {
        masaluCalendarRichMenuRequest_('richmenu/'+state.pending.id,token,'delete',null,404);
        delete state.pending;
        properties.setProperty(key,JSON.stringify(state));
      }
      if (!state.pending) {
        if (!template) {
          const source = JSON.parse(UrlFetchApp.fetch(assets+page.personal).getContentText());
          source.pixels = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(source.pixels))).getBytes();
          if (source.pixels.length !== 1250*422 || source.palette.length !== 768) throw new Error('個人月曆底圖格式錯誤');
          template = source;
        }
        const entries = masaluPersonalMenuEntries_(page,rows);
        const image = masaluPersonalMenuImage_(template,entries);
        const menu = JSON.parse(JSON.stringify(page.menu));
        menu.name = 'MaSaLu 個人排班 '+page.start;
        menu.size = {width:1250,height:422};
        menu.areas = entries.map(({bounds,action})=>({bounds,action}));
        menu.areas.push({bounds:{x:14,y:372,width:1222,height:42},action:page.menu.areas.at(-1).action});
        const created = masaluCalendarRichMenuRequest_('richmenu',token,'post',menu);
        const id = JSON.parse(created.getContentText()).richMenuId;
        try {
          masaluUploadCalendarRichMenu_(id,token,'image/png',image);
        } catch (error) {
          masaluCalendarRichMenuRequest_('richmenu/'+id,token,'delete',null,404);
          throw error;
        }
        state.pending = {id,signature};
        properties.setProperty(key,JSON.stringify(state));
      }
      // Retry an interrupted link only while its schedule and owner are current.
      masaluCalendarRichMenuRequest_(userPath+'/'+state.pending.id,token,'post');
      if (requestedUid && !masaluCalendarMenuLinked_(userPath,state.pending.id,token)) return {status:'failed',entries:ownEntries};
      const next = {...state.pending,retired:state.id || ''};
      properties.setProperty(key,JSON.stringify(next));
      if (next.retired) {
        masaluCalendarRichMenuRequest_('richmenu/'+next.retired,token,'delete',null,404);
        delete next.retired;
        properties.setProperty(key,JSON.stringify(next));
      }
    } catch (error) {
      failures++;
      // LINE endpoint errors may contain a UID. Keep only the status in logs.
      console.error('個人月曆選單更新失敗 '+(String(error.message).match(/HTTP \d+/)?.[0] || ''));
    }
  }
  if (requestedUid) {
    if (failures) return {status:'failed',entries:ownEntries};
    if (!targets.has(requestedUid)) return {status:'unbound',entries:[]};
    return {status:ownEntries.length ? 'linked' : 'empty',entries:ownEntries};
  }
  return failures;
}

function masaluCalendarMenuLinked_(userPath,id,token) {
  const response = masaluCalendarRichMenuRequest_(userPath,token,'get',null,404);
  return response.getResponseCode() === 200 && JSON.parse(response.getContentText()).richMenuId === id;
}

function masaluPersonalMenuEntries_(page, rows) {
  const days = new Map();
  for (const [date,site,confirmed] of rows) {
    if (!days.has(date)) days.set(date,{sites:new Set(),confirmed:true});
    const day = days.get(date);
    day.sites.add(site);
    day.confirmed = day.confirmed && confirmed;
  }
  if (!days.size) return [];
  const perColumn = days.size > 6 ? Math.ceil(days.size/2) : days.size;
  const cardWidth = days.size > 6 ? 605 : 1222;
  const step = Math.floor(300 / perColumn);
  return [...days].sort(([a],[b])=>a.localeCompare(b)).map(([date,day],index)=>{
    const offset = (Date.parse(date)-Date.parse(page.start))/86400000;
    return {date,sites:[...day.sites],confirmed:day.confirmed,action:page.menu.areas[offset].action,
      bounds:{x:14+Math.floor(index/perColumn)*617,y:66+(index%perColumn)*step,width:cardWidth,height:step-5}};
  });
}

function masaluPersonalMenuImage_(template, entries) {
  const width = 1250, height = 422;
  const pixels = Uint8Array.from(template.pixels);
  const glyphs = Object.fromEntries(Object.entries(template.glyphs).map(([label,value])=>[label,Utilities.base64Decode(value)]));
  const drawText = (label,x,y,color,scale) => {
    const mask = glyphs[label];
    for (let line=0; line<26; line++) {
      for (let column=0; column<60; column++) {
        if (!mask[line*60+column]) continue;
        for (let dy=0; dy<scale; dy++) {
          const start = (y+line*scale+dy)*width+x+column*scale;
          pixels.fill(color,start,start+scale);
        }
      }
    }
  };
  for (const {date,sites,confirmed,bounds} of entries) {
    const {x,y,width:cardWidth,height:cardHeight} = bounds;
    for (let line=y; line<y+cardHeight; line++) pixels.fill(252,line*width+x,line*width+x+cardWidth);
    const scale = cardHeight>=58 ? 2 : 1;
    const top = y+Math.floor((cardHeight-26*scale)/2);
    const dateLabel = date.slice(5).replace('-', '/');
    [...dateLabel].forEach((char,index)=>drawText(char,x+16+index*13*scale,top,254,scale));
    sites.forEach((site,index)=>drawText(site,x+110*scale+index*52*scale,top,253,scale));
    drawText(confirmed?'已確認':'待回覆',x+cardWidth-80*scale,top,confirmed?254:255,scale);
  }
  return masaluIndexedPng_(width,height,pixels,template.palette);
}

// Indexed PNG with uncompressed DEFLATE blocks stays below LINE's 1 MB limit.
function masaluIndexedPng_(width,height,pixels,palette) {
  const out = [137,80,78,71,13,10,26,10];
  const u32 = value => [(value>>>24)&255,(value>>>16)&255,(value>>>8)&255,value&255];
  const append = (target,bytes) => { for (const byte of bytes) target.push(byte); };
  const crcTable = Array.from({length:256},(_,value)=>{
    for (let i=0;i<8;i++) value = (value>>>1) ^ (value&1 ? 0xedb88320 : 0);
    return value;
  });
  function chunk(name,bytes) {
    append(out,u32(bytes.length));
    let crc = 0xffffffff;
    for (const byte of [...name].map(letter=>letter.charCodeAt(0))) {
      out.push(byte);
      crc = (crc>>>8) ^ crcTable[(crc^byte)&255];
    }
    for (const byte of bytes) {
      out.push(byte);
      crc = (crc>>>8) ^ crcTable[(crc^byte)&255];
    }
    append(out,u32(crc ^ 0xffffffff));
  }
  chunk('IHDR',[...u32(width),...u32(height),8,3,0,0,0]);
  chunk('PLTE',palette);
  const scanlines = new Uint8Array((width+1)*height);
  for (let y=0;y<height;y++) scanlines.set(pixels.subarray(y*width,(y+1)*width),y*(width+1)+1);
  const compressed = [0x78,0x01];
  let a=1,b=0;
  for (let start=0;start<scanlines.length;start+=65535) {
    const end = Math.min(start+65535,scanlines.length);
    const length = end-start;
    compressed.push(end===scanlines.length?1:0,length&255,length>>>8,(~length)&255,((~length)>>>8)&255);
    for (let i=start;i<end;i++) {
      const byte = scanlines[i];
      compressed.push(byte);
      a = (a+byte)%65521;
      b = (b+a)%65521;
    }
  }
  append(compressed,u32((b<<16)|a));
  chunk('IDAT',compressed);
  chunk('IEND',[]);
  return out.map(byte=>byte>127?byte-256:byte);
}
