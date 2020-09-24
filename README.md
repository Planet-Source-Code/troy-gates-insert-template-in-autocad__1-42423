<div align="center">

## Insert Template in AutoCAD


</div>

### Description

SrA John Spence asked for some examples for VBA with AutoCAD, so here is the first.

This is an example for creating a new drawing in AutoCAD and inserting a template. This is coded for VBA but is easily ported to VB 6. I can show that also if someone wants it.
 
### More Info
 
Must have AutoCAD

This needs to be in its own module or form and not ThisDrawing. Easiest is to add a new form and a button and put it in the button click event.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Troy Gates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/troy-gates.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/troy-gates-insert-template-in-autocad__1-42423/archive/master.zip)





### Source Code

```
Dim acadApp As AcadApplication 'reference to the AutoCAD application
  Dim acadDocs As AcadDocuments  'reference to the AutoCAD Documents collection
  Dim acadDoc As AcadDocument  'reference to a Document in the Collection
  Dim acadBlock As AcadBlockReference 'Block reference
  Dim strTemplate As String  'path to template file
  Dim dblInsertPt(2) As Double  'array with insert points (X,Y,Z)
  strTemplate = "S:\D+Acad\D+Templates\APF-Floor Plan.dwt" 'change to the path of the template you want
  Set acadApp = ThisDrawing.Application  'connect to AutoCAD application
  Set acadDocs = acadApp.Documents  'get the Documents collection
  Set acadDoc = acadDocs.Add 'create an empty document
  'set the inseration points to 0,0,0
  dblInsertPt(0) = 0# 'X
  dblInsertPt(1) = 0# 'Y
  dblInsertPt(2) = 0# 'Z
  'Insert the template with no XYZ scale and no rotation
  Set acadBlock = acadDoc.ModelSpace.InsertBlock(dblInsertPt, strTemplate, 1, 1, 1, 0)
  acadBlock.Explode  'explode the template
  acadApp.ZoomExtents 'zoom to the extents
  'clear objects from memory
  Set acadBlock = Nothing
  Set acadDoc = Nothing
  Set acadDocs = Nothing
  Set acadApp = Nothing
```

