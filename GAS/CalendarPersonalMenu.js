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
  const prefix = 'MASALU_PERSONAL_MENU_';
  let ownEntries = [];
  const saved = properties.getProperties();
  let template;
  let failures = 0;
  let failureStage = '';
  const users = requestedUid ? [requestedUid] :
    new Set([...targets.keys(), ...Object.keys(saved).filter(key => key.startsWith(prefix)).map(key => key.slice(prefix.length))]);
  for (const uid of users) {
    if (!validLineUid_(uid)) continue;
    const key = prefix + uid;
    const employeeId = targets.get(uid);
    let stage = '讀取個人班表';
    try {
      const entries = employeeId ? masaluPersonalRoster_(page,calendar,data,shifts,employeeId) : [];
      if (requestedUid) ownEntries = entries;
      const state = JSON.parse(saved[key] || '{}');
      stage = '檢查既有選單';
      if (state.retired) {
        masaluCalendarRichMenuRequest_('richmenu/'+state.retired,token,'delete',null,404);
        delete state.retired;
        properties.setProperty(key, JSON.stringify(state));
      }
      const signature = entries.length ? Utilities.base64EncodeWebSafe(
        Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(['grid', page.personal, entries]))
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
          stage = '讀取與解壓底圖';
          if (!page.personal.startsWith('personal-grid-')) throw new Error('個人月曆新版底圖尚未同步');
          const source = JSON.parse(UrlFetchApp.fetch(assets+page.personal).getContentText());
          source.pixels = masaluUnpackMenuAsset_(Utilities.base64Decode(source.pixels));
          if (source.pixels.length !== 1250*422 || source.palette.length !== 768) throw new Error('個人月曆底圖格式錯誤');
          stage = '讀取與解壓字型';
          source.rosterGlyphs = masaluUnpackMenuAsset_(UrlFetchApp.fetch(assets+'roster-glyphs.gz').getContent());
          if (source.rosterGlyphs.length !== (95+0xa000-0x3400)*36) throw new Error('班表字型格式錯誤');
          template = source;
        }
        stage = '繪製個人班表';
        const image = masaluPersonalMenuImage_(template,entries);
        const menu = JSON.parse(JSON.stringify(page.menu));
        menu.name = 'MaSaLu 個人排班 '+page.start;
        menu.size = {width:1250,height:422};
        menu.areas = menu.areas.map(({bounds,action})=>({action,bounds:{
          x:Math.floor(bounds.x/2),y:Math.floor(bounds.y/2),
          width:Math.floor((bounds.x+bounds.width)/2)-Math.floor(bounds.x/2),
          height:Math.floor((bounds.y+bounds.height)/2)-Math.floor(bounds.y/2)
        }}));
        stage = '建立 LINE 選單';
        const created = masaluCalendarRichMenuRequest_('richmenu',token,'post',menu);
        const id = JSON.parse(created.getContentText()).richMenuId;
        try {
          stage = '上傳個人班表圖片';
          masaluUploadCalendarRichMenu_(id,token,'image/png',image);
        } catch (error) {
          try {
            masaluCalendarRichMenuRequest_('richmenu/'+id,token,'delete',null,404);
          } catch (_) {
            console.error('未完成的個人選單清理失敗');
          }
          throw error;
        }
        state.pending = {id,signature};
        properties.setProperty(key,JSON.stringify(state));
      }
      // Retry an interrupted link only while its schedule and owner are current.
      stage = '套用到 LINE 帳號';
      masaluCalendarRichMenuRequest_(userPath+'/'+state.pending.id,token,'post');
      if (requestedUid && !masaluCalendarMenuLinked_(userPath,state.pending.id,token)) {
        return {status:'failed',entries:ownEntries,failureStage:'LINE 尚未回報選單已套用'};
      }
      const next = {...state.pending,retired:state.id || ''};
      properties.setProperty(key,JSON.stringify(next));
      if (next.retired) {
        try {
          masaluCalendarRichMenuRequest_('richmenu/'+next.retired,token,'delete',null,404);
          delete next.retired;
          properties.setProperty(key,JSON.stringify(next));
        } catch (_) {
          console.error('舊個人選單待清理');
        }
      }
      if (requestedUid) return {status:'linked',entries:ownEntries};
    } catch (error) {
      failures++;
      // LINE endpoint errors may contain a UID. Keep only the status in logs.
      const status = String(error.message).match(/HTTP \d+/)?.[0];
      failureStage = status ? stage+'（'+status+'）' : stage;
      console.error('個人月曆選單更新失敗：'+failureStage);
    }
  }
  if (requestedUid) {
    if (failures) return {status:'failed',entries:ownEntries,failureStage};
    if (!targets.has(requestedUid)) return {status:'unbound',entries:[]};
    return {status:ownEntries.length ? 'linked' : 'empty',entries:ownEntries};
  }
  return failures;
}

function masaluUnpackMenuAsset_(bytes) {
  return Utilities.ungzip(Utilities.newBlob(bytes,'application/x-gzip','menu.gz')).getBytes();
}

function masaluPersonalRoster_(page, calendar, data, shifts, employeeId) {
  const schedules = calendar.schedules.filter(row => row.date >= page.start && row.date <= page.end && shifts.has(row.shiftId));
  const visible = {schedules,staff:calendar.staff,shifts:calendar.shifts};
  masaluEmployeeCalendar_(visible,data,employeeId);
  const staff = new Map(visible.staff.map(person=>[person.id,person]));
  const byDate = new Map();
  const ownRows = [];
  for (const row of visible.schedules) {
    const location = shifts.get(row.shiftId);
    if (row.staffId === employeeId) ownRows.push([row.date,row.replyStatus === '同意']);
    if (!byDate.has(row.date)) byDate.set(row.date,new Map());
    const locations = byDate.get(row.date);
    if (!locations.has(location)) locations.set(location,new Map());
    const person = staff.get(row.staffId);
    if (person) locations.get(location).set(person.id,{
      name:person.name,mark:Array.from(String(person.short || person.name).trim())[0] || '?',
      color:person.color,
      own:person.id === employeeId
    });
  }
  return masaluPersonalMenuEntries_(page,ownRows).map(entry=>({...entry,
    groups:[...byDate.get(entry.date)].map(([location,people])=>({location,people:[...people.values()]}))
  }));
}

function masaluCalendarMenuLinked_(userPath,id,token) {
  const response = masaluCalendarRichMenuRequest_(userPath,token,'get',null,404);
  return response.getResponseCode() === 200 && JSON.parse(response.getContentText()).richMenuId === id;
}

function masaluPersonalMenuEntries_(page, rows) {
  const days = new Map();
  for (const [date,confirmed] of rows) {
    days.set(date,(days.get(date) ?? true) && confirmed);
  }
  return [...days].sort(([a],[b])=>a.localeCompare(b)).map(([date,confirmed])=>{
    const offset = (Date.parse(date)-Date.parse(page.start))/86400000;
    return {date,confirmed,bounds:page.menu.areas[offset].bounds};
  });
}

function masaluPersonalMenuImage_(template, entries) {
  const width = 1250, height = 422;
  const pixels = Uint8Array.from(template.pixels);
  const glyphs = Object.fromEntries(Object.entries(template.glyphs).map(([label,value])=>[label,Utilities.base64Decode(value)]));
  const drawText = (label,x,y,color) => {
    const mask = glyphs[label];
    const maskWidth = mask.length === 72*32 ? 72 : 60;
    for (let line=0; line<mask.length/maskWidth; line++) {
      for (let column=0; column<maskWidth; column++) {
        if (mask[line*maskWidth+column]) pixels[(y+line)*width+x+column] = color;
      }
    }
  };
  const drawGlyph = (char,x,y,color,scale) => {
    const code = char.codePointAt(0);
    const index = code>=32 && code<127 ? code-32 : code>=0x3400 && code<0xa000 ? 95+code-0x3400 : 31;
    for (let line=0; line<18; line++) {
      for (let column=0; column<16; column++) {
        if (!(template.rosterGlyphs[index*36+line*2+(column>>3)] & (128>>(column%8)))) continue;
        for (let targetY=Math.floor(line*scale); targetY<Math.floor((line+1)*scale); targetY++) {
          const start=(y+targetY)*width+x;
          pixels.fill(color,start+Math.floor(column*scale),start+Math.floor((column+1)*scale));
        }
      }
    }
  };
  const palette = [...template.palette];
  const personColors = new Map();
  const tileColor = color => {
    const hex = /^#[0-9a-f]{6}$/i.test(color || '') ? color.toUpperCase() : '#475569';
    if (personColors.has(hex)) return personColors.get(hex);
    if (personColors.size === 16) return {fill:254,ink:252};
    const rgb = [1,3,5].map(offset=>parseInt(hex.slice(offset,offset+2),16));
    const fill = 236 + personColors.size;
    palette.splice(fill*3,3,...rgb);
    const ink = (rgb[0]*299 + rgb[1]*587 + rgb[2]*114) >= 150000 ? 252 : 253;
    const colors = {fill,ink};
    personColors.set(hex,colors);
    return colors;
  };
  const rectangle = (x,y,w,h,color) => {
    for (let line=y;line<y+h;line++) pixels.fill(color,line*width+x,line*width+x+w);
  };
  for (const {groups,confirmed,bounds} of entries) {
    const x = Math.floor(bounds.x/2);
    const y = Math.floor(bounds.y/2);
    rectangle(x+93,y+2,73,32,252);
    drawText(confirmed?'已確認':'待回覆',x+94,y+2,confirmed?254:255);
    const shown = groups.length > 3 ? groups.slice(0,2) : groups;
    const roomy = shown.length < 3 && groups.length === shown.length;
    const tile = roomy ? 32 : 24;
    const scale = roomy ? 1.75 : 1.375;
    const labelWidth = roomy ? 62 : 51;
    const capacity = Math.floor((165-labelWidth)/(tile+2));
    shown.forEach(({location,people},index)=>{
      const top = y+34+index*(tile+2);
      rectangle(x+4,top,161,tile,252);
      const label = Array.from(location).slice(0,2);
      label.forEach((char,i)=>drawGlyph(char,x+5+i*Math.floor(16*scale),top,253,scale));
      const limit = people.length > capacity ? capacity-1 : capacity;
      const ordered = [...people].sort((a,b)=>Number(b.own)-Number(a.own));
      ordered.slice(0,limit).forEach((person,i)=>{
        const left = x+labelWidth+i*(tile+2);
        const {fill,ink} = tileColor(person.color);
        rectangle(left,top,tile,tile,person.own?254:fill);
        drawGlyph(person.mark,left+Math.floor((tile-16*scale)/2),top,person.own?252:ink,scale);
      });
      if (people.length>limit) drawGlyph('+',x+labelWidth+limit*(tile+2),top,255,scale);
    });
    if (groups.length>shown.length) drawGlyph('+',x+6,y+85,255,1);
  }
  return masaluIndexedPng_(width,height,pixels,palette);
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
