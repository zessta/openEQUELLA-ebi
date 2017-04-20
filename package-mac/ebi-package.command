cd "$(dirname "$0")"

python setup.py py2app --iconfile ebibig.icns  --plist Info.plist

mv dist/ebi.app/Contents/MacOS/ebi.command dist/ebi.app/Contents/MacOS/ebi

cp source/ebi/ebibig.ico dist/ebi.app/Contents/Resources/lib/.
cp source/ebi/fileopen.png dist/ebi.app/Contents/Resources/lib/.
cp source/ebi/filesave.png dist/ebi.app/Contents/Resources/lib/.
cp source/ebi/gtk-help.png dist/ebi.app/Contents/Resources/lib/.
cp source/ebi/gtk-stop.png dist/ebi.app/Contents/Resources/lib/.
cp source/ebi/options.png dist/ebi.app/Contents/Resources/lib/.
cp source/ebi/pause.png dist/ebi.app/Content
frjikjhjkfdf021s/Resources/lib/.
cp ebi.command dist/ebi.app/Contents/MacOS/.

cd dist
hdiutil create -srcfolder ebi.app ebi.dmg
hdiutil internet-enable -yes ebi.dmg