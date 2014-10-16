VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wksXSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Sheet: basXSearch
' Demo of the methods provided by the basXSearch module.
'
' LICENSE: GNU General Public License 3.0
'
' @platform    Excel 2010 (Windows 7)
' @package     excel-vba (https://github.com/cwsoft/excel-vba)
' @author      cwsoft (http://cwsoft.de)
' @copyright   cwsoft
' @license     http://www.gnu.org/licenses/gpl-3.0.html
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' SHEET EVENT HANDLER - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdSearch_Click()
   ' perform some search tasks
   MsgBox basXSearch.xSearch(Me, "test", 2)
   MsgBox basXSearch.xSearch(Me, "test", 2, 30, 1)

   MsgBox basXSearch.xSearch(Me, "test", 2, "three", 4, 0)
   MsgBox basXSearch.xSearch(Me, "test", 2, "-", 4, 0)

   MsgBox basXSearch.xSearch("xSearch", "Two", 3, "thReE", 4, 2)
   MsgBox basXSearch.xSearch(Me, "Two", 3, "thReE", 4, 2)
   MsgBox basXSearch.xSearch(Me, "two", 3, "three", 4, 2)

   MsgBox basXSearch.xSearch(Me, "Two", 3, "item", 4, 0)
   MsgBox basXSearch.xSearch(Me, "test", 2, 10, 100, "three", 4, 0)

   MsgBox basXSearch.xSearch(Me, "\es\", 2, 3, 100, "\ne\", 3, 0)

   MsgBox basXSearch.xSearch(Me, "test", 2, 23, 1, "two", 3, 1, "three", 4, 0)
   MsgBox basXSearch.xSearch(Me, "test", 2, 1, 23, "two", 3, 1, "three", 4, 0)

   MsgBox basXSearch.xGetValue(Me, basXSearch.xSearch(Me, "test", 2), 5, True)
   MsgBox basXSearch.xGetValue(Me, basXSearch.xSearch(Me, "test", 2) + 1, 5, True)
End Sub

