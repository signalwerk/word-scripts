## Word-Scripts

Script to copy image-formatting from InDesign to Word

### `copyPicFormat`
copy a formatting of a picture form InDesign (`.jsx`) to Word (`.vbs`)

### `multipleTemplates`
Generate multiple Templates based on the open document. Copy text to replace in each new Template;

Clipboard data (tab separated):

| First       | Second     | DEL_PIC      |
| ----------- | ---------- | ------------ |
| Content A1  | Content B1 | hello\|world |
| Content A2  | Content B2 | hello        |

Replaces in the Doc `{{First}} {{Second}}` with `Content A1 Content B1`.  
Removes picture (`Shape` not `InlineShapes`) that have the Alt-Text `hello` (Line 1 & 2) and `world` (Line 1).  
The resulting filename of the new template will be `{{First}}.dotm`.


### Customize the Word User Interface (Ribbon)
Example:
![Ribbon Example](./doc/img/logo-ein-aus.png)
* [Handmade](https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file)
* With [Custom UI Editor](http://gregmaxey.com/word_tip_pages/customize_ribbon_main.html)
* [Download Custom UI Editor](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2006/05/26/customuieditor.aspx)

### References
See also:
* [Trista DPI – Batch processing of raster images in Adobe InDesign documents](https://github.com/flekst/Trista-DPI)

### Status
This is only a set of test-scripts. use at your own risk. :)

## Images in Word
If you have a Image in Illustrator you can import it with `.emf` (Bezier/Spline-Suppoert) or `.wmf` (only Polyline) to Word. 
