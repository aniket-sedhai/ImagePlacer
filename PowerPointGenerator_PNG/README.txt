PowerPointGenerator.exe produces .pptx files by appending images founding within the "images" folder 
that exist in the following structure:

		directory structure: "Output 20xxxxxx//[Any Folder]//images//pre_6_l.png"

PowerPointGenerator.exe goes through all "Any Folder" in the folder "Output 20xxxxxx" and finds all the images named "pre_6_l.png"
in the folder "images". It pulls all the images that were found into a powerpoint file and adds those images into one slide per image.

User Instructions:

***IMPORTANT***
DO NOT DELETE THE FILES PowerPointGenerator or default.pptx. The application cannot be used without either one of these files.
Always make sure there is a default powerpoint file "default.pptx" along with "PowerPointGenerator.exe" in the same folder.

1. Run PowerPointGenerator.exe
   - A dialog box asking for you to choose original folder should appear in few seconds
2. Choose the desired folder that contains the folder "Output 20xxxxxx" with the images inside it. Do not choose
   "Output 20xxxxxx" itself.
   - Once you select that original folder, a new dialog box should appear asking for destination folder.
3. Choose any destination folder where you would like the PowerPoint Presentation file to end up.

If the folder "Output 20xxxxxx" contains folders that contain folder "images", which in turn contains the image file "pre_6_l.png",
all those images will be added to the presentation file and the presentation file will open up.

#Note: Make sure you close the presentation file that opened up before running "PowerPointGenerator.exe" again.
