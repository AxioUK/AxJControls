# AxJControls
Collections of Controls ComboBox an List with Columns (based on originals combo &amp; list of J. Elihu)

## 📦 axJList 

ListBox multicolumna oculto, se "engancha" a un textbox mediante su _hWnd_ el cual se debe pasar con la función _.Init {hWnd}_
Una vez "enganchado" con _.ShowList_ e _.HideList_ mostramos u ocultamos el List.
Para recuperar los valores de las columnas de un item se usa el evento _ItemClick()_

## 📦 axJColCombo

ComboBox multicolumna, se muestran todas las columnas en el list, pero solo se envía un dato al textbox. El dato a enviar al textbox se define con la propiedad 
_.ColumnInBox_, para recuperar los valores de las columnas de un item se usa el evento _ItemClick()_

## 📦 axJCombo

ComboBox simple, se pueden ingresar hasta 4 "columnas" con el _.AddItem_, de las que solo se muestra una, segun se indique en la propiedad _.ColumInList_, 
la misma columna en el List será la que se muestra en el textBox.  
En el evento _ItemClick()_ se puede recuperar el ListIndex del Item y los valores de sus 4 "columnas" como strings.

### Pre-requisitos 📋

Nada en Particular....


### Instalación 🔧



## Autores ✒️

* **J. Elihú** - *Trabajo Inicial, desarrollador de los usercontrol originales JList y JCombo* - 
* **D. Rojas** - *Modificaciones y adaptaciones AxJList, AxJCombo y AxJColCombo* -

También puedes mirar la lista de todos los [contribuyentes](https://github.com/your/project/contributors) quíenes han participado en este proyecto.
