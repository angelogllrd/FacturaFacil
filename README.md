# FacturaFacil
GUI para automatizar la creación de facturas en la web de AFIP/ARCA usando Selenium.

> [!NOTE]
> Si bien el programa está hecho para un flujo de trabajo específico, el código es adaptable a otros casos.

![Sin-título](https://github.com/user-attachments/assets/04dc84de-2c83-4bb1-93ac-7a3fb56dea10)

## Cómo probarlo

1. Clonar repositorio: 
    
    ```
    git clone https://github.com/angelogllrd/FacturaFacil.git && cd FacturaFacil
    ```

2. Instalar requerimientos: 

    ```
    pip install -r requirements.txt
    ```

    o bien ejecutar directamente esto:

    ```
    pip install PyQt5 selenium pyperclip
    ```
3. Ejecutar main.py:
    ```
    python main.py
    ```
El programa (así como está) necesita:
1. Google Chrome instalado.

2. Un archivo `credenciales.json` en la carpeta del proyecto, que contenga:
    ```
    {
        "cuit": string_de_cuit_válido_en_afip,
        "clave": string_de_clave_de_afip,
        "cuit_receptor": cuit_de_un_receptor_válido
    }
    ```
    que se usará para ingresar a la cuenta de AFIP, y seleccionar el destinatario de la factura.

3. `chromedriver.exe` en la carpeta del proyecto, compatible con la versión instalada de Chrome. Se descarga de [acá](https://developer.chrome.com/docs/chromedriver/downloads) (versiones de Chrome anteriores a la 115) o [acá](https://googlechromelabs.github.io/chrome-for-testing/#stable) (versión 115 o posteriores). 

    También se puede usar **webdriver-manager**, que se encarga de instalar y actualizar el driver automáticamente:
    
    * Instalar **webdriver-manager**:
        
        ```
        pip install webdriver-manager
        ```
    * En `afip_automation.py` debe modificarse lo siguiente:

        ```
        # Descomentar
        from webdriver_manager.chrome import ChromeDriverManager

        # Comentar
        # service = Service(CHROME_DRIVER_PATH)

        # Descomentar
	    service = Service(ChromeDriverManager().install())
        ```
    
4. Que se copie al portapapeles, desde Excel o Google Sheets, una porción de tabla de 7 u 8 columnas de la forma:
   
    <table>
      <tr><td>9232</td><td>27/1/2025</td><td>SA</td><td>TINT</td><td>descripción del trabajo</td><td>Hecho</td><td>$12.345</td><td>$20.500</td></tr>
      <tr><td>9157</td><td>27/1/2025</td><td>SA</td><td>TINT</td><td>descripción del trabajo</td><td>Hecho</td><td>$67.898</td><td>$392.200</td></tr>
      <tr><td>9278</td><td>27/1/2025</td><td>SA</td><td>TINT</td><td>descripción del trabajo</td><td>Hecho</td><td>$76.543</td><td>$33.3428</td></tr>
    </table>

    donde:
   
    |col1|col2|col3|col4|col5|col6|col7|col8|
    |---|---|---|---|---|---|---|---|
    |String de números decimales (0-9)| Fecha con año de 4 digitos al inicio o al final y separadores "-", "/", o "." | String alfabético (a-zA-Z) con longitud no mayor a 10 | Idem columna anterior | No se controla | No se controla | String de dinero sin "$" ni "." con números decimales (0-9) y como máximo una "," | (OPCIONAL) Idem columna anterior

## Pasar de tabla a estructura de datos

> [!NOTE]
> Esto explica lo que hace la función `formatClipboard()` de `clipboard_utils.py`, que toma una porción de planilla copiado como texto plano, y la transformaa a una estructura de datos más manejable.

Si copio esta porción de tabla de Google Sheets (o Excel):

![img1](https://github.com/user-attachments/assets/aa1d496b-0d68-4289-bbfa-02b342522f00)

y la pego en el editor, se ve así:

```
9038	29/1/2025	Lorem ipsum dolor sit amet, consectetur adipiscing elit	$101.157
9233	27/1/2025	"Nulla ut lorem a orci pulvinar ornare euismod at eros.
Aliquam commodo dapibus
Pellentesque auctor vestibulum"	$55.124
9221	5/2/2025	Sed luctus est sit amet justo vestibulum, vel rutrum erat vulputate.	$172.066
9158	28/12/2024	"Aenean non odio accumsan, ornare turpis et.
Praesent pretium facilisis consequat"	$721.364
```

donde las columnas están separadas por una tabulación `\t`, y cada fila por un salto de línea `\n`. Sin embargo, el texto de las celdas que tienen texto en varias líneas **se pega en líneas diferentes**, cuando deberían pertenecer a la misma. Esto dificulta separar el texto en las filas originales para hacerlo manejable en el código.

Una forma de corregir esto es reemplazar los `\n` dentro de estas celdas por un espacio y mantener los `\n` que realmente separan filas, pero **¿cómo diferenciar uno de otro?**. 

Con `repr()` podemos ver el texto "crudo" copiado al portapapeles:

```
import pyperclip

cb = pyperclip.paste()
print(repr(cb))
````
que devuelve:

```
'9038\t29/1/2025\tLorem ipsum dolor sit amet, consectetur adipiscing elit\t$101.157\r\n9233\t27/1/2025\t"Nulla ut lorem a orci pulvinar ornare euismod at eros.\nAliquam commodo dapibus\nPellentesque auctor vestibulum"\t$55.124\r\n9221\t5/2/2025\tSed luctus est sit amet justo vestibulum, vel rutrum erat vulputate.\t$172.066\r\n9158\t28/12/2024\t"Aenean non odio accumsan, ornare turpis et.\nPraesent pretium facilisis consequat"\t$721.364'
```

Para diferenciar los saltos de línea dentro de celdas con los que separan filas podemos:
* **Reemplazar los `\n` que no son precedidos por un `\r`:** El `\r` es un "carriage return" (retorno de carro) y aparece porque el texto copiado usa `\r\n` (retorno de carro + nueva línea) como separador de líneas, lo cual es común en sistemas Windows. Solamente los saltos de línea de nuevas filas tienen un `\r` antes.
* **Reemplazar todos los \n que están entre `\t"` y `"\t`:** Las celdas con texto en varias lineas se pegan con comillas dobles en el inicio y final del texto. Como este puede tener, a su vez, comillas dobles dentro, puedo diferenciar las iniciales y finales porque están pegadas a un `\t` que marca la separación con la columna anterior y posterior.

Para hacer lo segundo, uso lo siguiente:

```
cb = re.sub(r'(\t".*?"\t)', lambda m: m.group(1).replace('\n', ' '), cb, flags=re.DOTALL)
```
* La expresión regular `r'(\t".*?"\t)'` y re.DOTALL:
    * `\t"` busca el tabulador seguido de una comilla doble.
    * `.*?` captura cualquier texto, incluyendo saltos de línea (gracias a `re.DOTALL`), de manera no codiciosa (non-greedy o lazy). 
    
        **¿Por qué se usa non-greedy matching?** Si no usáramos el `?` y ejecutamos lo siguiente:
        
        ```
        cb = re.sub(r'(\t".*"\t)', 'hola', cb, flags=re.DOTALL)
        ```

        se devuelve lo siguiente:
        
        ```
        9038	29/1/2025	Lorem ipsum dolor sit amet, consectetur adipiscing elit	$101.157
        9233	27/1/2025hola$721.364
        ```
        es decir, reemplaza todo desde el primer `\t"` hasta el último `"\t` de **toda** la tabla. En cambio si usamos la búsqueda no codiciosa (poniendo `?`) se devuelve esto:

        ```
        9038	29/1/2025	Lorem ipsum dolor sit amet, consectetur adipiscing elit	$101.157
        9233	27/1/2025hola$55.124
        9221	5/2/2025	Sed luctus est sit amet justo vestibulum, vel rutrum erat vulputate.	$172.066
        9158	28/12/2024hola$721.364
        ```
        es decir, reemplaza desde el primer `\t"` hasta el `"\t` próximo más cercano en cada coincidencia.

    * `"\t` busca la comilla doble de cierre seguida de un tabulador.
* `lambda m: m.group(1).replace('\n', ' ')`: Si usara `r'\1'.replace('\n', ' ')` estaría intentando usar `replace()` antes de que `re.sub()` haga su trabajo, es decir, se haría un `replace('\n', ' ')` sobre `'\1'`. La solución es usar una función `lambda`, donde:
    * `m` es el match encontrado por re.sub().
    * `m.group(1)` obtiene el texto del grupo capturado.
    * `.replace('\n', ' ')` reemplaza los saltos de línea en ese texto.

Por último, hay que quitar las comillas dobles del texto de aquellas celdas que tenían saltos de línea. Para eso uso regexes con **lookbehind** y **lookahead**, donde "miro" inmediatamente antes y después de las comillas dobles para reemplazar aquellas que tienen un `\t` pegado:

```
cb = re.sub(r'(?<=\t)"', '', cb)
cb = re.sub(r'"(?=\t)', '', cb)
```

De esta forma, `formatClipboard()` retorna una tabla (lista de listas) de la siguiente forma:

```
[['9038', '29/1/2025', 'Lorem ipsum dolor sit amet, consectetur adipiscing elit', '$101.157'],
 ['9233', '27/1/2025', 'Nulla ut lorem a orci pulvinar ornare euismod at eros. Aliquam commodo dapibus Pellentesque auctor vestibulum', '$55.124'],
 ['9221', '5/2/2025', 'Sed luctus est sit amet justo vestibulum, vel rutrum erat vulputate.', '$172.066'],
 ['9158', '28/12/2024', 'Aenean non odio accumsan, ornare turpis et. Praesent pretium facilisis consequat', '$721.364']]
```

