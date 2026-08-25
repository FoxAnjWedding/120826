const SHEET_ID = "1es7PorZHtDHGO6-f78-RJYs5X8DEChvMKA6930RvRC4";
const ADMIN_KEY = "060897rR.";

function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const guests = getOrCreate_(ss, "Guests", ["id","group","name","primaryName","seats","mobile","dietary","message","createdAt"]);
  const responses = getOrCreate_(ss, "Responses", ["timestamp","guestId","group","name","status","mobile","dietary","message"]);
  if (guests.getLastRow() === 1) {
    guests.getRange(2,1,4,9).setValues([
      ["G001","Sample Family","Juan Dela Cruz","Juan Dela Cruz",2,"","","",new Date()],
      ["G002","Sample Family","Maria Dela Cruz","Juan Dela Cruz",2,"","","",new Date()],
      ["G003","Sample Family","Pedro Dela Cruz","Juan Dela Cruz",2,"","","",new Date()],
      ["G004","Sample Couple","Ana Santos","Ana Santos",2,"","","",new Date()]
    ]);
  }
  return "Setup complete. Replace the sample guests in the Guests sheet with your real guest list.";
}
function getOrCreate_(ss,name,headers){
  let sh=ss.getSheetByName(name)||ss.insertSheet(name);
  if(sh.getLastRow()===0) sh.appendRow(headers);
  return sh;
}
function doPost(e){
  try{
    const p=JSON.parse(e.postData.contents||"{}");
    if(p.action==="findGuest") return json_(findGuest_(p.name));
    if(p.action==="submitRSVP") return json_(submit_(p));
    if(p.action==="adminData") return json_(admin_(p));
    return json_({ok:false,message:"Unknown action"});
  }catch(err){return json_({ok:false,message:err.message})}
}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function sheet_(n){return SpreadsheetApp.openById(SHEET_ID).getSheetByName(n)}
function rows_(sh){const v=sh.getDataRange().getValues();const h=v.shift();return v.filter(r=>r.join("")!=="").map(r=>Object.fromEntries(h.map((x,i)=>[x,r[i]])))}
function findGuest_(name){
  const gs=rows_(sheet_("Guests")); const q=name.toLowerCase().trim();
  const matches=gs.filter(g=>String(g.name).toLowerCase().trim()===q);
  if(!matches.length) return {ok:false,message:"We couldn't find that name. Please check the spelling."};
  const group=matches[0].group, people=gs.filter(g=>g.group===group);
  const resp=rows_(sheet_("Responses")).filter(r=>r.group===group);
  const latest={}; resp.forEach(r=>latest[r.guestId]=r);
  return {ok:true,household:{id:group,primaryName:matches[0].primaryName||matches[0].name,seats:Math.max(...people.map(x=>Number(x.seats)||1)),mobile:resp[resp.length-1]?.mobile||"",dietary:resp[resp.length-1]?.dietary||"",message:resp[resp.length-1]?.message||"",guests:people.map(g=>({id:g.id,name:g.name,status:latest[g.id]?.status||""}))}};
}
function submit_(p){
  const gs=rows_(sheet_("Guests")); const allowed=gs.filter(g=>g.group===p.householdId);
  const sh=sheet_("Responses"); const now=new Date();
  Object.keys(p.answers||{}).forEach(id=>{
    const g=allowed.find(x=>x.id===id); if(!g) throw new Error("Invalid guest.");
    sh.appendRow([now,id,g.group,g.name,p.answers[id],p.mobile||"",p.dietary||"",p.message||""]);
  });
  return {ok:true};
}
function admin_(p){
  if(p.adminKey!==ADMIN_KEY) return {ok:false,message:"Invalid admin key."};
  const gs=rows_(sheet_("Guests")), rs=rows_(sheet_("Responses")), latest={};
  rs.forEach(r=>latest[r.guestId]=r);
  return {ok:true,rows:gs.map(g=>({name:g.name,group:g.group,status:latest[g.id]?.status||"Pending",mobile:latest[g.id]?.mobile||"",dietary:latest[g.id]?.dietary||"",message:latest[g.id]?.message||"",updated:latest[g.id]?.timestamp||""}))};
}