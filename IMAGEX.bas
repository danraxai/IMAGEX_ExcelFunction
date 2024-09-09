' Función en inglés: IMAGEX
Public Function IMAGEX(celdaOUrl As Variant, celdaDestino As Range) As String
    IMAGEX = InsertarImagen(celdaOUrl, celdaDestino)
End Function

' Función en español: IMAGENX
Public Function IMAGENX(celdaOUrl As Variant, celdaDestino As Range) As String
    IMAGENX = InsertarImagen(celdaOUrl, celdaDestino)
End Function

' Función compartida para insertar imagen
Public Function InsertarImagen(celdaOUrl As Variant, celdaDestino As Range) As String
    Dim urlImagen As String
    Dim imagenDescargada As Picture
    Dim imagen As Object

    ' Determinar si la entrada es una celda o una URL directa
    If TypeName(celdaOUrl) = "Range" Then
        urlImagen = celdaOUrl.Value
    Else
        urlImagen = celdaOUrl
    End If

    ' Eliminar imágenes existentes en la celda de destino
    For Each imagen In ActiveSheet.Pictures
        If Not Intersect(imagen.TopLeftCell, celdaDestino) Is Nothing Then
            imagen.Delete
        End If
    Next imagen

    If urlImagen = "" Then
        InsertarImagen = "URL no provista"
        Exit Function
    End If

    On Error GoTo ErrHandler
    Set imagenDescargada = ActiveSheet.Pictures.Insert(urlImagen)

    With imagenDescargada
        .ShapeRange.LockAspectRatio = msoTrue
        .Left = celdaDestino.Left
        .Top = celdaDestino.Top
        .Width = celdaDestino.Width
        .Height = celdaDestino.Height
    End With

    ' Asegurar que la imagen se mueva y cambie de tamaño con la celda
    imagenDescargada.Placement = xlMoveAndSize

    InsertarImagen = "Success"
    Exit Function

ErrHandler:
    InsertarImagen = "URL Invalida o error en conexion de Red"
End Function

' Sub para registrar las funciones IMAGEX y IMAGENX
Sub RegistrarFuncionesIMAGEX_IMAGENX()
    ' Registro de IMAGEX
    Application.MacroOptions _
        Macro:="IMAGEX", _
        Description:="Inserts an image into the worksheet from a specified URL or from a cell containing a URL.", _
        Category:="Custom", _
        ArgumentDescriptions:=Array("Cell or URL of the image", "Destination cell where the image will be inserted")

    ' Registro de IMAGENX (en español)
    Application.MacroOptions _
        Macro:="IMAGENX", _
        Description:="Inserta una imagen en la hoja de cálculo desde una URL especificada o desde una celda que contiene una URL.", _
        Category:="Personalizadas", _
        ArgumentDescriptions:=Array("Celda o URL de la imagen", "Celda de destino donde se insertará la imagen")
End Sub
