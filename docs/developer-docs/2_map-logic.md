---
id: map-logic
title: Логика работы графической части листа "Карта"
---


# Общая информация

Ниже описана логика работы графической части, созданной на листе **"Карта"** в Excel. Она основана на стандартном функционале Excel для добавления фигур и соединительных линий. Скрипт отвечает за построение визуальной схемы объектов (исток, объекты, трубы, стоки) и их автоматическое соединение, а также за назначение цветов и макросов для обработки событий (например, двойной клик по объекту).

## Создание фигур (CreateShape)

Основная логика создания фигур описана в процедуре `CreateShape`. Она принимает ряд параметров, указывающих, какую фигуру нужно создать и куда её разместить:

- **modelType** – модель объекта (Исток, Труба, Объект, Сток).  
- **objectName** – название объекта.  
- **GroupType** – тип объекта (Исток, НП, ГП, НПС, БКНС и т. д.).  
- **i** – номер строки объекта на листе **"МТБ"**.  
- **cellLeft**, **cellTop** – координаты по оси X и Y, где будет вставлена фигура.  
- **endCellLeft**, **endCellTop** – координаты окончания трубы (соединительной линии), если это Труба.

Ниже приведён код процедуры:
```basic
Sub CreateShape(ByVal modelType As String, ByVal objectName As String, GroupType As String,_
                ByVal i As Integer, Optional cellLeft As Double = 100, Optional cellTop As Double = 100,_
                Optional endCellLeft As Double = 150, Optional endCellTop As Double = 150)
    ' Создаем необходимую фигуру в соотвествии с объектом(исток - овал, объект - прямоугольник, труба - линия, сток - треугольник)
    Dim ws As Worksheet
    Dim isLine As Boolean
    
    Dim shapeGroup As shape
    Dim shape As shape
    Dim shapeText As shape
    Dim shapeLine As shape
    Dim startShape As shape
    Dim endShape As shape
    
    Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
    Dim midX As Integer, midY As Integer
    Dim isTransit As String
    
    isTransit = ""

    Dim targetCell As Range
    Sheets("Карта").Activate
    Set ws = ActiveSheet
    Set targetCell = ws.Range("M3") ' Ячейка рядом с которой будет создан объект
    ' Выбор типа фигуры
    isLine = False
    macroName = "ShowUserForm"
```
### Проверка существующей фигуры
Перед созданием новой фигуры проверяется, есть ли уже фигура с таким именем (заданным номером строки i) на листе "Карта". Если такая фигура существует, создание прерывается:

```basic

    For Each shp In ws.Shapes
        If shp.name = CStr(i) Then
            MsgBox "Фигура уже на доске установлена" & objectName
            shapeExists = True
            Exit Sub
        End If
    Next shp
```
### Определение типа создаваемой фигуры
В зависимости от modelType определяется, какой тип фигуры нужно нарисовать. Ниже приведён пример соответствия типа модели и фигуры:

    - Исток – овал
    - Объект – прямоугольник (или овал, если транзит)
    - Труба – соединительная линия
    - Сток – прямоугольник со скруглёнными углами (вместо треугольника в примере)
На рисунке ниже представлены примеры типов фигур:

![Типы объектов](./img/objects.PNG "Типы объектов")


```basic
    Select Case modelType
        Case "Исток"
            ' Создание овала
            Set shape = ws.Shapes.AddShape(msoShapeOval, cellLeft, cellTop, 100, 50)
            shape.TextFrame.Characters.Text = objectName
            
        Case "Объект"
            
            isTransit = Sheets("МТБ").Cells(i + 2, 15).value
            ' Создание прямоугольника
            If isTransit = "Транзит" Then
                ' Set shape = ws.Shapes.AddShape(msoShapeRectangle, cellLeft, cellTop, 100, 50)
                Set shape = ws.Shapes.AddShape(msoShapeOval, cellLeft, cellTop, 80, 45)
            Else
                Set shape = ws.Shapes.AddShape(msoShapeRectangle, cellLeft, cellTop, 100, 50)
            End If
            
            shape.TextFrame.Characters.Text = objectName
        Case "Труба"
            ' Создание линии
            ' Set shapeLine = ws.Shapes.AddLine(cellLeft, cellTop, endCellLeft, endCellTop)
            Set shapeLine = ws.Shapes.AddConnector(msoConnectorElbow, cellLeft, cellTop, endCellLeft, endCellTop)
            shapeLine.line.EndArrowheadStyle = msoArrowheadTriangle ' делаем линию - стрелкой
            shapeLine.name = i
            ColorShapeByGroupName GroupType, shapeLine
            ' Привязка макроса к линии
            shapeLine.OnAction = macroName
            shapeLine.line.Weight = 2
            ' Проверяем есть рядом с созданной линией объекты(фигуры) и соединияем если есть
            CheckAndConnectNearestShapes ws, shapeLine, cellLeft, cellTop, endCellLeft, endCellTop
            
            x1 = shapeLine.Left
            y1 = shapeLine.Top
            x2 = x1 + shapeLine.Width
            y2 = y1 + shapeLine.Height
            
            midX = (x1 + x2) / 2
            midY = (y1 + y2) / 2
            
            shapeLine.OnAction = macroName

        Case "Сток"
            ' Создание треугольника
            Set shape = ws.Shapes.AddShape(msoShapeRoundedRectangle, cellLeft, cellTop, 100, 50)
            shape.TextFrame.Characters.Text = objectName
        Case Else
            MsgBox "Неизвестный тип модели: " & modelType, vbExclamation
            Exit Sub
    End Select
```
### Дополнительные настройки и назначение макросов
Для вновь созданной фигуры задаются:
    -  Имя фигуры (shape.Name) – номер строки из листа "МТБ".
    - Макрос при двойном щелчке (shape.OnAction).
    - Заливка и контур (ColorShapeByGroupName).
    - Дополнительный текст или атрибуты (например, AlternativeText).
```basic
    ' Установка названия для фигуры
    If Not shape Is Nothing And isLine = False Then
        If isTransit = "Транзит" Then
            shape.Fill.ForeColor.RGB = RGB(128, 128, 128) ' Средний серый
        Else
            ColorShapeByGroupName GroupType, shape
        End If
        
        shape.name = i
        shape.OnAction = macroName
        shape.AlternativeText = objectName & "&" & modelType & "&" & GroupType
    End If
End Sub
```

### Автоматическое соединение фигур (CheckAndConnectNearestShapes)
Для удобства автоматического соединения труб с ближайшими фигурами используется процедура CheckAndConnectNearestShapes. Она размещена в модуле CreateShapeModule.

```basic
Sub CheckAndConnectNearestShapes(ws As Worksheet, lineShape As shape, cellLeft As Double, cellTop As Double, endCellLeft As Double, endCellTop As Double)
    Dim startCell As Range
    Dim endCell As Range
    Dim startShape As shape
    Dim endShape As shape

    ' Шаг 2: Проверяем, есть ли рядом фигура с началом линии
    Set startShape = FindNearestShape(ws, cellLeft, cellTop)

    ' Шаг 3: Проверяем, есть ли рядом фигура с концом линии
    Set endShape = FindNearestShape(ws, endCellLeft, endCellTop)
    
    ' Шаг 5: Если нашли фигуры рядом, привязываем линию к ним
    If Not startShape Is Nothing Then
        ConnectLineToShape lineShape, startShape, True ' Привязываем начало линии
    End If

    If Not endShape Is Nothing Then
        ConnectLineToShape lineShape, endShape, False ' Привязываем конец линии
    End If

End Sub
```
### Поиск ближайшей фигуры (FindNearestShape)
Функция FindNearestShape ищет фигуру в пределах радиуса minDistance (по умолчанию 150 пунктов), которая поддерживает точки подключения (фигура не должна быть линией или свободной формой).
```basic
' Функция, которая ищет ближайшую фигуру к указанной позиции, поддерживающую соединители
Function FindNearestShape(ws As Worksheet, cellLeft As Double, cellTop As Double) As shape
    Dim shp As shape
    Dim minDistance As Double
    Dim closestShape As shape
    Dim distance As Double

    minDistance = 150 ' Максимальное расстояние до фигуры (можно настроить)

    For Each shp In ws.Shapes
        ' Проверяем, что фигура поддерживает точки подключения
        If shp.AutoShapeType <> msoShapeLine And shp.AutoShapeType <> msoShapeFreeform And shp.Type <> msoLine Then
            If shp.ConnectionSiteCount > 0 Then
                ' Вычисляем расстояние до фигуры
                distance = CalculateDistance(cellLeft, cellTop, shp.Left + shp.Width / 2, shp.Top + shp.Height / 2)
                If distance < minDistance Then
                    Set closestShape = shp
                    minDistance = distance
                End If
            End If
        End If
    Next shp

    Set FindNearestShape = closestShape
End Function


' Функция для расчета расстояния между двумя точками
Function CalculateDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    CalculateDistance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

' Функция для привязки линии к фигуре
Sub ConnectLineToShape(line As shape, shape As shape, isStart As Boolean)
    If isStart Then
        ' Привязываем начало линии к фигуре
        line.ConnectorFormat.BeginConnect shape, 3
    Else
        ' Привязываем конец линии к фигуре
        line.ConnectorFormat.EndConnect shape, 1
    End If
End Sub
```


### Раскрашивание фигур (ColorShapeByGroupName)
    Процедура ColorShapeByGroupName, расположенная в модуле Colorizer, задаёт цвет заливки и контура фигур в зависимости от их типа. Для некоторых типов (например, линии) цвет применяется только к контуру, а для других — и к заливке, и к контуру.

```basic

Public Sub ColorShapeByGroupName(groupName As String, shp As shape)
    ' Цветовое распределение объектов
    Dim colorValue As Long
    
    Select Case groupName
        Case "Исток"
            colorValue = RGB(204, 153, 255) ' Светло-сиреневый
        Case "Сток"
            colorValue = RGB(102, 51, 153) ' Темно-сиреневый
        Case "НП", "НПС", "ПСП"
            colorValue = vbBlack  ' Черный
        Case "ЦПС", "УПСВ", "ДНС", "МФНС", "ПТ"
            colorValue = RGB(131, 60, 11)  ' Коричневый
        Case "БКНС", "ВД"
            colorValue = RGB(68, 114, 196)  ' Синий
        Case "ГП", "УКПГ", "ДКС", "ГКС"
            colorValue = RGB(255, 192, 0)  ' Темно-желтый
        Case "ПП"
            colorValue = RGB(237, 125, 49)  ' Оранжевый
        Case "Транзит"
            colorValue = RGB(128, 128, 128) ' Средний серый
        Case Else
            ' Если тип не распознан, устанавливаем цвет по умолчанию (серый)
            colorValue = vbGray
    End Select
    
    ' Применяем цвет к фигуре
    On Error Resume Next
    If groupName = "ГП" Or groupName = "НП" Or groupName = "ПП" Or groupName = "ВД" Or groupName = "ПТ" Then
        ' Если фигура является линией, применяем цвет только к контуру
        shp.line.ForeColor.RGB = colorValue
        shp.line.Weight = 3 ' Толщина линии
    Else
        ' Если фигура не линия, применяем цвет заливки и контура
        shp.Fill.ForeColor.RGB = colorValue
        shp.line.ForeColor.RGB = colorValue
        shp.line.Weight = 5 ' Увеличиваем толщину линии
        
        ' Проверяем, является ли цвет заливки чёрным
        If colorValue = vbBlack Or colorValue = RGB(102, 51, 153) Then
            shp.TextFrame.Characters.Font.Color = vbWhite
        ElseIf colorValue = RGB(204, 153, 255) Or colorValue = RGB(255, 192, 0) Then  ' Темно-желтый
            shp.TextFrame.Characters.Font.Color = vbBlack
        End If
    End If
    On Error GoTo 0
End Sub
```

## Построение Линии трубопроводов