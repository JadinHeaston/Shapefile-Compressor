# shapefile-compressor
The basic function is that the script compresses files of a similar name (excluding extensions) within a directory (Does not recursively look through sub-directories). Primarily created for shapefiles.

Eventually, I may add an additional check for the specific shapefile extensions. - Although, I have been using C++ more often recently and the speed gains from using that may prove to be worth the effort of remaking this simple script.


This script was created to automatically zip shapefiles into a single compressed folder.

It will compress all items with the same name (different extensions) into a single zip folder. It will NOT delete any of the original data.
It will prompt for a folder to begin compressing and will immediately go to work. A completion dialog will appear when finished.

Notes
  - XML files will act weird. The test data (and use case) this was created for meant that the file name for .xml files were "*.shp.xml"
	- There is an optional log that can be enabled by changing "logKeeper" to be a 1 instead of a 0, near the top of the script.
*	- There is 700ms pause when creating zip files before other things continue, due to the script attempting to copy files before the zip 
	file has enough time to be created fully.
		- This sleep timer can be modified at the bottom of the "NewZip" function near the bottom.
*	- There is another 130ms sleep timer under the "CopyToZip" function, designed to give the file time to copy before moving on.
		- This is about as low as you should go, although, larger files might require a longer time.
