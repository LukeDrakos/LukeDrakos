const CFB = require('cfb'); const fs = require('fs');
const [,, inp, out, dir, mod, vbap] = process.argv;
const cfb = CFB.read(fs.readFileSync(inp), {type:'buffer'});
for (const p of cfb.FullPaths.slice()) { if (/__SRP_\d+$/.test(p)) CFB.utils.cfb_del(cfb, p.replace(/^Root Entry/,'')); }
const rep = (name, file) => { const ok = CFB.utils.cfb_del(cfb, name); CFB.utils.cfb_add(cfb, name, fs.readFileSync(file)); };
rep('/VBA/dir', dir); rep('/VBA/modEI32', mod); rep('/VBA/_VBA_PROJECT', vbap);
fs.writeFileSync(out, CFB.write(cfb, {type:'buffer', fileType:'cfb'}));
console.log(cfb.FullPaths.filter(p=>!p.endsWith('/')).join('\n'));
