const fs=require('node:fs'),project=process.env.SONAR_PROJECT_KEY;
if(!project)throw Error('SONAR_PROJECT_KEY is required.');
const pr=process.env.AUDIT_PR||'',branch=process.env.AUDIT_BRANCH||'main',scope=pr?{pullRequest:pr}:{branch},headers={Accept:'application/json'};
if(process.env.SONAR_TOKEN)headers.Authorization='Basic '+Buffer.from(process.env.SONAR_TOKEN+':').toString('base64');
fs.mkdirSync('audit-results',{recursive:true});
async function get(endpoint,params={}){
const url=new URL('/api/'+endpoint,'https://sonarcloud.io');for(const[k,v]of Object.entries(params))url.searchParams.set(k,v);
const r=await fetch(url,{headers,redirect:'error',signal:AbortSignal.timeout(25000)});
if(!r.ok)throw Error(endpoint+': HTTP '+r.status);
const d=await r.json();if(d.errors)throw Error(endpoint+': '+d.errors.map(e=>e.msg).join('; '));return d;
}
async function pages(endpoint,params,field){
let items=[],p=1;while(true){const d=await get(endpoint,{...params,p,ps:500});items.push(...(d[field]||[]));const total=d.paging?.total??d.total??items.length;
if(items.length>=total)return{items,total,complete:true};if(p*500>=10000)return{items,total,complete:false};p++;}
}
(async()=>{
const report={project,scope,requestedSourceSha:process.env.AUDIT_SOURCE_SHA||'',mode:'SonarCloud automatic analysis',runtimeMeasured:false,status:'PENDING'};let lastError='';
for(let attempt=0;attempt<24;attempt++){
try{
const ghUrl='https://api.github.com/repos/'+process.env.GITHUB_REPOSITORY+'/commits/'+report.requestedSourceSha+'/check-runs';
const response=await fetch(ghUrl,{headers:{Accept:'application/vnd.github+json',Authorization:'Bearer '+process.env.GH_TOKEN},redirect:'error',signal:AbortSignal.timeout(25000)});
if(!response.ok)throw Error('GitHub analysis verification: HTTP '+response.status);
const checks=await response.json(),check=(checks.check_runs||[]).find(c=>c.name==='SonarCloud Code Analysis'&&c.head_sha===report.requestedSourceSha&&c.status==='completed'&&/Quality Gate/i.test(c.output?.title||''));
if(!check)throw Error('SonarCloud has not completed analysis of the exact requested source commit.');
report.sonarCheck={name:check.name,headSha:check.head_sha,conclusion:check.conclusion,url:check.details_url};
const d=await get('components/tree' ,{component:project,qualifiers:'FIL',...scope,ps:500}),paths=(d.components||[]).map(c=>c.path);
if(paths.includes('sonar-src/Code.js')&&paths.includes('sonar-src/Index.js')){report.analyzedPaths=paths;report.status='ANALYZED';break;}
lastError='Selected analysis does not yet contain both latest analyzer mirrors.';}catch(e){lastError=e.message;}
if(attempt<23)await new Promise(r=>setTimeout(r,10000));}
if(report.status==='ANALYZED'){
for(const[key,endpoint,params,field]of[['issues','issues/search',{componentKeys:project,resolved:'false',...scope},'issues'],['hotspots','hotspots/search',{projectKey:project,...scope},'hotspots']]){
try{report[key]=await pages(endpoint,params,field);}catch(e){report[key]={error:e.message};}}
for(const[key,endpoint,params]of[['qualityGate','qualitygates/project_status',{projectKey:project,...scope}],['measures','measures/component',{component:project,...scope,metricKeys:'ncloc,bugs,vulnerabilities,code_smells,duplicated_lines_density,cognitive_complexity,security_rating,reliability_rating,sqale_rating'}]]){
try{report[key]=await get(endpoint,params);}catch(e){report[key]={error:e.message};}}
}else report.blocker=lastError;
report.sourceManifest=JSON.parse(fs.readFileSync('sonar-src/source-manifest.json','utf8'));
fs.writeFileSync('audit-results/sonarcloud-report.json',JSON.stringify(report,null,2));
const issues=report.issues?.items||[],severity=issues.reduce((a,x)=>{const k=x.severity||'UNSPECIFIED';a[k]=(a[k]||0)+1;return a;},{});
const summary='## SonarCloud audit\n\nStatus: '+report.status+'\n\nScope: '+JSON.stringify(scope)+'\n\nRequested source SHA: '+report.requestedSourceSha+'\n\nBoth latest source mirrors present: '+(report.status==='ANALYZED'?'yes':'not verified')+'\n\nOpen issues retrieved: '+issues.length+'\n\nSeverity: '+JSON.stringify(severity)+'\n\nQuality Gate: '+(report.qualityGate?.projectStatus?.status||'not available')+'\n\n'+(report.blocker?'Blocker: '+report.blocker+'\n\n':'')+'Static analysis does not measure login duration, Sheets latency, or concurrent-user capacity.\n';
fs.writeFileSync('audit-results/sonarcloud-summary.md',summary);if(process.env.GITHUB_STEP_SUMMARY)fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY,summary);console.log(summary);
if(report.status!=='ANALYZED')process.exitCode=1;
})().catch(e=>{console.error(e.message);process.exitCode=1;});
