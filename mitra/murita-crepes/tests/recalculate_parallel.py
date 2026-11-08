"""Independent LibreOffice profiles and content-keyed temporary QA caches."""
import concurrent.futures,datetime,hashlib,os,shutil,subprocess,tempfile,zipfile
from pathlib import Path

CACHE=Path('/tmp/mufrita-pos-qa-cache')

def fingerprint(path):
    h=hashlib.sha256(datetime.date.today().isoformat().encode())
    with zipfile.ZipFile(path) as z:
        for name in sorted(z.namelist()):
            if name=='docProps/core.xml':continue
            h.update(name.encode());h.update(z.read(name))
    return h.hexdigest()

def calculate(cmd,**kwargs):
    CACHE.mkdir(exist_ok=True);out=Path(cmd[5]);tasks=[]
    for name in cmd[6:]:
        src=Path(name);cached=CACHE/(fingerprint(src)+'.xlsx')
        if cached.exists():shutil.copy2(cached,out/src.name)
        else:tasks.append((src,cached))
    def run_one(task):
        src,cached=task
        with tempfile.TemporaryDirectory(prefix='mufrita-office-') as profile:
            cache_dir=Path(profile)/'cache';cache_dir.mkdir()
            args=[cmd[0],'-env:UserInstallation='+Path(profile).as_uri(),*cmd[1:6],str(src)]
            result=subprocess.run(args,env={**os.environ,'XDG_CACHE_HOME':str(cache_dir)},**kwargs)
            assert result.returncode==0,f'Office exit {result.returncode}: '+result.stdout+'\n'+result.stderr
            assert (out/src.name).exists(),result.stdout+'\n'+result.stderr
            with zipfile.ZipFile(out/src.name) as z:assert z.testzip() is None
            shutil.copy2(out/src.name,cached)
    # A fresh process per file releases office memory between scenarios.
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        futures=[executor.submit(run_one,task) for task in tasks]
        for future in futures:future.result()
    return subprocess.CompletedProcess(cmd,0,stdout='Content-verified scenario recalculation complete.',stderr='')
