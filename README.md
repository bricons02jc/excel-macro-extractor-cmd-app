# Extractor de macros de Excel - *"Excel Macro Extractor"*

Aplicación de cmd de Windows para extraer archivos `.vba` de archivos `.xlsm`.


## Uso:

~~~ sh
$ ExcelMacroExtrator.exe file targetdir [--copy-xlsm]

# ejemplo
$ ExcelMacroExtrator.exe "C:\Dev\Archivo.xlsm" "C:\Dev\Archivo-Fuente" --copy-xlsm
~~~

La opción `--copy-xlsm` también copia el archivo de Excel en el `targetpath`. 

## Motivación
Con esta herramienta puedo extraer el código VBA a un `targetdir` (directorio-objetivo) y comparar versiones de ese directorio con `git`.
#

>*"I develop a fair amount of Excel-VBA-Applications, and want to track the VBA-modules via `git`. The Excel-file itself is a binary-file, and therefore not really `diff`-able. With this tool I can extract the VBA-code to a `targetdir` and track that directory with `git`."*
- Dominik Aumayr - [aumayr](https://github.com/aumayr "Perfil GitHub")
