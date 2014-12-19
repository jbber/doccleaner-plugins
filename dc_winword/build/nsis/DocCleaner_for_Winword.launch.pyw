#!python3.4-32
import sys, os
scriptdir, script = os.path.split(__file__)
pkgdir = os.path.join(scriptdir, 'pkgs')
sys.path.insert(0, pkgdir)
os.environ['PYTHONPATH'] = pkgdir + os.pathsep + os.environ.get('PYTHONPATH', '')

appdata = os.environ.get('APPDATA', None)
def excepthook(etype, value, tb):
    "Write unhandled exceptions to a file rather than exiting silently."
    import traceback
    with open(os.path.join(appdata, script+'.log'), 'w') as f:
        traceback.print_exception(etype, value, tb, file=f)

if appdata:
    sys.excepthook = excepthook



if __name__ == '__main__':
    from wordaddin import main
    main()
