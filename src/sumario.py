import subprocess
import sys
import time
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException
from pathlib import Path


# Atualiza o sumário do .docx usando o LibreOffice.
# Roda no GitHub Actions, com o Python do sistema (pacote python3-uno).

arquivo = Path(sys.argv[1]).resolve()
endereco = 'socket,host=localhost,port=2002;urp;'

libreoffice = subprocess.Popen(
    ['soffice', '--headless', '--invisible', f'--accept={endereco}']
)
contexto = uno.getComponentContext()
resolvedor = contexto.ServiceManager.createInstanceWithContext(
    'com.sun.star.bridge.UnoUrlResolver',
    contexto,
)
for _ in range(100):
    try:
        remoto = resolvedor.resolve(
            f'uno:{endereco}StarOffice.ComponentContext'
        )
        break
    except NoConnectException:
        time.sleep(0.3)

desktop = remoto.ServiceManager.createInstanceWithContext(
    'com.sun.star.frame.Desktop',
    remoto,
)
documento = desktop.loadComponentFromURL(
    arquivo.as_uri(),
    '_blank',
    0,
    (PropertyValue(Name='Hidden', Value=True),),
)
indices = documento.getDocumentIndexes()
for i in range(indices.getCount()):
    indices.getByIndex(i).update()
documento.storeToURL(
    arquivo.as_uri(),
    (PropertyValue(Name='FilterName', Value='MS Word 2007 XML'),),
)
documento.close(True)
desktop.terminate()
libreoffice.wait()
print(f'Sumário atualizado: {arquivo.name}')
