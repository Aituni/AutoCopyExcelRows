Nota: los comandos van sin las " ".

(solo la primera vez)
Antes de usarlo abre el terminal y ejecuta cada uno de estos comandos: (es para instalar las librerias necesarias)

"pip install glob"
"pip install tqdm"
"pip install pandas"
"pip install xlrd"


Para usar el programa tienes que abrir el terminal en la dirección o carpeta en la que esté el archivo ejecutable "main.py".
Una vez tengas el terminal abierto ejecuta el siguiente comando:

"Python .\main.py"

La primera vez te creará la carpeta de archivos si es que no la encuentra. Mete todos los excels (haz una copia porseaca) en la carpeta "Archivos"
y vuelve a ejecutar exactamente el mismo comando. Ahora sí, se deberia de haber ejecutado por completo. En cada ejecucion, se borraran las bases de datos (si es que habia alguna)
y creara nuevas y actualizadas, asi que si quieres mantener alguna haz copia pega en otra carpeta antes de usar el comando. Despues, 
automaticamente se pondra a rellenar las tablas, los excels rellenados se guardaran todos en la carpeta "Rellenados", 
si existia de antes algun archivo que tenga el mismo nombre lo sobrescribira.

El comando genera dos bases de datos, una con las filas completadas que ha encontrado ("base de datos.xlsx") y otra con las filas incompletas que ha encontrado ("base de datos PENDIENTES.xlsx"),
en ambos casos, los datasets contienen los valores encontrados a lo largo de todos los excels de archivos, y no hay filas repetidas.
si actualizas la "base de datos.xlsx", y metes nuevas filas rellenadas de las que habia pendiente, puedes ejecutar el siguiente comando:

"Python .\main.py 2"

Con este comando, se rellenaran todos los excels de archivos, pero haciendo uso de la "base de datos.xlsx" que haya hecha.