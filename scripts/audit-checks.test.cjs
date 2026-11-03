const assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs');
const source=fs.readFileSync(__dirname+'/sonar-report.cjs','utf8');
async function fixture(mode){
const written={},calls=[],fakeFS={mkdirSync(){},readFileSync(){return '{"version":"2.9.0"}';},writeFileSync(p,v){written[p]=v;},appendFileSync(){}};
const proc={env:{SONAR_PROJECT_KEY:'PROJECT',AUDIT_PR:'7',AUDIT_BRANCH:'feature',AUDIT_SOURCE_SHA:'HEAD',GITHUB_REPOSITORY:'owner/repo',GH_TOKEN:'TEST_GH_SECRET',SONAR_TOKEN:'TEST_SONAR_SECRET'},exitCode:0};
const fetch=async(input)=>{const url=new URL(input);calls.push(url);if(url.hostname==='api.github.com')return {ok:true,json:async()=>({check_runs:[{name:'SonarCloud Code Analysis',head_sha:mode==='stale'?'OLD':'HEAD',status:'completed',conclusion:'failure',output:{title:'Quality Gate failed'},details_url:'https://sonarcloud.io/dashboard?id=PROJECT&pullRequest=7'}]})};
assert.equal(url.searchParams.get('pullRequest'),'7');assert.equal(url.searchParams.has('branch'),false);
if(mode==='denied')return {ok:false,status:403};
let data={};if(url.pathname.endsWith('/tree'))data={components:[{path:'sonar-src/Index.js'},...(mode==='missing'?[]:[{path:'sonar-src/Code.js'}])]};
if(url.pathname.endsWith('/issues/search')){const p=Number(url.searchParams.get('p'));data={paging:{total:501},issues:Array.from({length:p===1?500:1},(_,i)=>({key:'K'+p+'-'+i,severity:'MAJOR'}))};}
if(url.pathname.endsWith('/hotspots/search'))data={paging:{total:0},hotspots:[]};
if(url.pathname.endsWith('/project_status'))data={projectStatus:{status:'ERROR'}};
return {ok:true,json:async()=>data};};
const promise=vm.runInNewContext(source,{require:n=>{assert.equal(n,'node:fs');return fakeFS;},process:proc,Buffer,URL,AbortSignal,fetch,setTimeout:fn=>{fn();return 0;},console:{log(){},error(m){throw Error(m);}}});await promise;
const out=JSON.parse(written['audit-results/sonarcloud-report.json']);assert.ok(!JSON.stringify(written).includes('TEST_GH_SECRET'));assert.ok(!JSON.stringify(written).includes('TEST_SONAR_SECRET'));return {out,calls,proc};
}
(async()=>{
let f=await fixture('valid');assert.equal(f.out.status,'ANALYZED');assert.equal(f.out.issues.items.length,501);assert.equal(f.out.qualityGate.projectStatus.status,'ERROR');assert.equal(f.out.sonarCheck.headSha,'HEAD');console.log('PASS current commit, PR scope, pagination, and failing Quality Gate remain accurate');
f=await fixture('stale');assert.equal(f.out.status,'PENDING');assert.equal(f.calls.some(u=>u.hostname==='sonarcloud.io'),false);assert.equal(f.proc.exitCode,1);console.log('PASS stale commit cannot be presented as the latest audit');
f=await fixture('missing');assert.equal(f.out.status,'PENDING');assert.match(f.out.blocker,/both latest/);console.log('PASS missing backend analysis cannot be reported as complete');
f=await fixture('denied');assert.equal(f.out.status,'PENDING');assert.match(f.out.blocker,/HTTP 403/);console.log('PASS unavailable SonarCloud report is explicit and secrets never enter artifacts');
})().catch(e=>{console.error(e);process.exitCode=1;});
