# %%
# Converts ppsx to pptx by patching content type
import os
import zipfile

def patch_ppsx(ppsx_filename):
    # Generates a new ppsx file which acts as a pptx file
    with zipfile.ZipFile(ppsx_filename) as z:
        with open('[Content_Types].xml', 'wb') as f:
            orignal_xml = z.read('[Content_Types].xml')
            patched_xml = orignal_xml.replace(b"slideshow.main", b"presentation.main")
            f.write(patched_xml)


    zin = zipfile.ZipFile (ppsx_filename, 'r')
    zout = zipfile.ZipFile (ppsx_filename[:-5]+'_patched.pptx', 'w')

    for item in zin.infolist():
        buffer = zin.read(item.filename)
        if (item.filename != '[Content_Types].xml'):
            zout.writestr(item, buffer)

    zout.write('[Content_Types].xml')
    zout.close()
    zin.close()
    os.remove("[Content_Types].xml")

# %%
if __name__ == "__main__":
    patch_ppsx('test.ppsx')