import zipfile, os

print "Zipping files..."
archive = zipfile.ZipFile("ebi.zip", 'w')
for root, dirs, files in os.walk('ebi'):
        for fileName in files:
                archive.write(os.path.join(root,fileName))
archive.close()
