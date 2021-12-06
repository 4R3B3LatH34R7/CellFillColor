# getFillColor
MS Excel UDF to get Cell fill color

## Background
As usual, in a facebook group, somebody asked for a formula to get a cell's fill color.\
As we all know, there is no such formula in Excel Worksheet UI.\
At first, I thought it was as easy as returning the output of Range.Interior.Color\
But it was more complicated than that as I found out.

First obstacle was the cells in question are part of a Pivot Table which was overcome by turning off the File->Options->Formulas->Use GetPivotData functions for Pivot table references.\
My Excel version is 2010 so YMMV.\

The second obstacle was that the Fill color of the cells in question are the result of a conditional formatting.\
Now this raises the bar to Range.DisplayFormat.Interior.Color OR Range.Cells.DisplayFormat.Interior.Color as per your preferences.

The next hurdle was that DisplayFormat if called from a UDF returned a #VALUE error.\
However, Jaafar Tribak's post at [MRExcel Forum](https://www.mrexcel.com/board/threads/using-displayformat-in-a-udf.1154593/) provided a workaround to this problem by the use of Evaluate function.\
Now I got the UDF going.

The last hurdle was trying to make the UDF compatible with multiple cell input and array output.\
For that I just loop through the range and use the above workaround and store the fill colors of cells in that range and return the array as variant.

## The VBA UDF Code
````VBA
Option Explicit

Function CellFillColor(target As Range, Optional returnFormat As String = "IDX") As Variant
Dim retArray()
Dim rowCounter As Long
Dim colCounter As Long
Dim colorValue As Long
    Application.Volatile
    If TypeName(target) = "Range" Then
        ReDim retArray(target.Rows.Count - 1, target.Columns.Count - 1)
        For rowCounter = 0 To target.Rows.Count - 1
            For colCounter = 0 To target.Columns.Count - 1
                colorValue = Evaluate("useDF(" & target.Cells(rowCounter + 1, colCounter + 1).Address & ")")
                Select Case UCase(returnFormat)
                    Case "RGB":
                                retArray(rowCounter, colCounter) = _
                                                                    Format((colorValue Mod 256), "00") & ", " & _
                                                                    Format(((colorValue \ 256) Mod 256), "00") & ", " & _
                                                                    Format((colorValue \ 65536), "00")
                    Case "HEX":
                                retArray(rowCounter, colCounter) = _
                                                                    "#" & _
                                                                    Format(Hex(colorValue Mod 256), "00") & _
                                                                    Format(Hex((colorValue \ 256) Mod 256), "00") & _
                                                                    Format(Hex((colorValue \ 65536)), "00")
                    Case "IDX": retArray(rowCounter, colCounter) = colorValue
                    Case Else: retArray(rowCounter, colCounter) = colorValue
                End Select
            Next colCounter
        Next rowCounter
        CellFillColor = retArray
    End If
End Function

Private Function useDF(ByVal target As Range) As Variant
    useDF = target.DisplayFormat.Interior.Color
End Function
````
The code above can be copied and pasted in a VBA code module and use as =CellFillColor(A1).\
There are 3 switches as arguments which will change the way the UDF returns the Fill color value of the cell. We can call the UDF as =CellFillColor(A1,returnFormat) with the following 3 possible values for returnFormat argument.\
1. "IDX" - default value and if set, the return will be a VBA color value.
2. "HEX" - a hexadecimal value in #FFFFFF format which can be used to change the color value from Excel UI in later versions of Excel. In Excel 2010, Hex input box is not found.
3. "RGB" - returns RGB values as (RED,GREEN,BLUE) for example as (255,102,133) etc. which can be used to change color of a cell using the Fill Color->More Colors->RGB.

In Excel 2010, the UDF can be entered as an array by pressing the Ctrl+Shift+Enter.\
However, an equavalent range of cells must be selected first before entering the array formula which is the norm I guess. The UDF needs to be entered as an array formula or else only the left top result will be returned.\
In Office365, this UDF will just spill the results over if Ctrl+Shift+Enter was not used to enter it.

## Releases
I will release the UDF in 3 forms.
1. UDF as function as copyable text above
2. .bas module
3. .xlsm file

## Future
Will try to fix bugs but for this, I am going to need the users feedback. Thank you.
