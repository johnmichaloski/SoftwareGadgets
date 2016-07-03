# Photo Indexer

This software primitively takes a collection of photos and generates an index of the photos.  If there are too many for one jpg index, it will generate a series of jpg indexes. Further, it will generate indexes for each folder under the starting folder, if the Recursive Index button is pushed.


After starting the program the user is greeted witht the following confusing dialog.

 ![Figure 1 Main Dialog](./images/MainDialog.JPG?raw=true)

- "Prepend" text box \- If text is inserted into the "Prepend" text box places text before the caption on each generated jpg index label. 

- "Index Label Off In Picture" checkbox button  \- will turn off labelling of jpg index

- "Generate sheet indexes" checkbox button \- unclear.

- "Generated Index" Push button \- 
- "Recuvseve Index" Push button \- will recursively search and generate jpg indexes for each subfolder. If labeling of indexes is chosen, each label will be the name of the subfolder, so it is helpful if the folder names are meaningful.

Importantly, one needs to browse for an input folder can pictures (and potentially subfolders containing more picture) and also an output folder for the indexes to be stored. Below shows the browse for folder dialog.
![Figure 2 Browse for Folders](./images/BrowseForFolder.JPG?raw=true)

Once the folders are selected, 

![Figure 3 After from and to folder selection](./images/FromToPictureFolders.JPG?raw=true)

Press the Recurse Index button, and wait.  The indexes should be generated, and a new tab should appear on the dialog box.

![Figure 4 Tab with indexes displays](./images/GeneratedIndexTab.JPG?raw=true)


Finally, you can open the index jpg in another browser.
![Figure 5 Sample Index](./images/ExampleIndexof2013 Nationals Game.JPG?raw=true)

This program needs polishing, however, if all you want are some indexes from a list of photos, it should do the trick. No promises though. 
