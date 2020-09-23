Notes

1. While printing the ScaleMode property of the container of List
view control must be in VbPixels unit.

2. This class works with listview control version 6.0 [SP 4]. 
    
3. Number of subitems must be equal.
--------

PosX and PosY: These require values in Twips unit.

LastRowPrinted: Turns True only if the last listitem in listview
control is printed.

SetLines: Set draw width and color for the lines. For printer you
can set non integer value such as 1.2, 1.4, 2.1 etc.

SetRows: This method set values to RowTo and RowFrom depending on
the value set to NumOfRowsPerPage.

NumOfRowsPerPage: Set this value if you are using SetRows method to 
print several pages.

CurrentX and CurrentY: The value returned is in Twips.

Rowheight: Value in Twips. This value should be tall enough to
fit Textheight and Picture height.

Opal
buna48@hotmail.com
http://geocities.com/opalraj/vb