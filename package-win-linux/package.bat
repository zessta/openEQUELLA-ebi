echo Y | RD /S dist
echo Y | del ebi
python setup.py py2exe
echo Y | RD /S ebi
echo Y | del ebi.zip
rename dist ebi
copy ebi.py ebi
copy MainFrame.py ebi
copy OptionsDialog.py ebi
copy Engine.py ebi
copy equellaclient41.py ebi
copy fileopen.png ebi
copy filesave.png ebi
copy gtk-stop.png ebi
copy gtk-help.png ebi
copy options.png ebi
copy pause.png ebi
copy ebibig.ico ebi
python package.py