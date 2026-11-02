const fs=require('node:fs'),vm=require('node:vm'),crypto=require('node:crypto'),path=require('node:path');
const backend=fs.readFileSync('Code.gs','utf8'),html=fs.readFileSync('Index.html','utf8');
const blocks=[...html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script\s*>/gi)].filter(m=>!/\bsrc\s*=/i.test(m[1]));
if(blocks.length!==1)throw Error('Expected exactly one inline application script. Review extraction before scanning.');
const block=blocks[0],start=block.index+block[0].indexOf('>')+1,frontend='\n'.repeat((html.slice(0,start).match(/\n/g)||[]).length)+block[2];
new vm.Script(backend,{filename:'Code.gs'});new vm.Script(frontend,{filename:'Index.html'});
const sha256=s=>crypto.createHash('sha256').update(s).digest('hex');
const generated={'sonar-src/Code.js':backend,'sonar-src/Index.js':frontend,'sonar-src/source-manifest.json':JSON.stringify({version:backend.match(/LDP_VERSION_\s*=\s*"([^"]+)"/)?.[1],sources:[
{source:'Code.gs',analyzerFile:'sonar-src/Code.js',sha256:sha256(backend),mapping:'Identical content and line numbers'},
{source:'Index.html',analyzerFile:'sonar-src/Index.js',sha256:sha256(html),analyzerSha256:sha256(frontend),mapping:'Inline application JS; original HTML line numbers preserved'}
]},null,2)+'\n'};
for(const[name,content]of Object.entries(generated)){
if(process.argv.includes('--check')){if(!fs.existsSync(name)||fs.readFileSync(name,'utf8')!==content)throw Error(name+' is stale. Run node scripts/prepare-sonar.cjs and commit analyzer mirrors.');}
else{fs.mkdirSync(path.dirname(name),{recursive:true});fs.writeFileSync(name,content);}
}
console.log('Source syntax and analyzer mirrors verified. Application files were not modified.');
