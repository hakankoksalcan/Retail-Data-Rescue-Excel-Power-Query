// Microsoft parser syntax checks; this does not execute M or refresh Excel.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const pq = require('@microsoft/powerquery-parser');
const root = path.resolve(__dirname, '..');
function files(dir) {
  return fs.readdirSync(dir, {withFileTypes:true}).flatMap(entry => {
    const full = path.join(dir, entry.name);
    return entry.isDirectory() ? files(full) : full.endsWith('.pq') ? [full] : [];
  });
}
(async () => {
  const results = [];
  for (const file of files(path.join(root, 'power_query')).sort()) {
    const result = await pq.TaskUtils.tryLexParse(pq.DefaultSettings, fs.readFileSync(file, 'utf8'));
    results.push({file:path.relative(root,file).split(path.sep).join('/'),sha256:crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex'),status:result.resultKind==='Ok'?'PASS':'FAIL',stage:result.stage,error:result.resultKind==='Ok'?null:String(result.error?.message || result.error)});
  }
  const out = {engine:'@microsoft/powerquery-parser 2.0.0',scope:'Syntax only; not semantic evaluation or Excel refresh',results};
  fs.writeFileSync(path.join(root,'validation','m-syntax.json'),JSON.stringify(out,null,2)+'\n');
  console.log(`${results.filter(r=>r.status==='PASS').length}/${results.length} M sources parsed successfully`);
  if (!results.length || results.some(r=>r.status!=='PASS')) process.exitCode=1;
})().catch(error => { console.error(error); process.exitCode=1; });
