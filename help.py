def help():
    """
    FolderGenerator v1.0 - 10/08/2020
    contact: buildertester

    Use Requirements
    Tool developed for Windows OS

    Operation Instructions
    1. Use a sheet within your Mapping XLS document to develop the target folder structure
    2. First row is reserved for headers and ignored by tool
    3. Column A is used to activate(1) or deactivate(0) folders and sub - folders for folder structure generation
    4. Column B should contain your main folder name whilst subsequent columns contain any desired number of sub - folder levels
    5. Tool will add numerical prefixes to folders according to the hierarchy you have developed therefore manual adding of numerical prefixes not required
    6. When ready save your Mapping XLS document
    7. Run FolderGenerator EXE
    8. Follow the prompts:
        'Select Mapping Document' will allow you to locate and select you Mapping XLS document(when the selection is successful the document path will be displayed under the 'Select Mapping Document' button)
        'Select Sheet' containing the target folder structure from the dropdown menu
        'Select Path' where you want to generate you new root directory (when the selection is successful the folder path will be displayed under the 'Select Path' button)
        Change 'InsertRootName' to a desired target root directory name
    9. After completing all of the inputs, click on Generate
    10. The indexed folder structure will appear on your target path

    Target Features for Future Releases
    1. Allow user to define hyperlinks in target folder structure e.g.Rlive standards / references to reduce size of root directories and avoid issues associated with duplicate files
    2. Allow user to define symbolic links in target folder structure e.g.link VSIM libraries in SVN Target folder to SVN MIL Plant Model folder to reduce size of SVN folder and avoid issues associated with duplicate files (might not be possible if SVN root is always in different locations on different user PCs, need investigation)
    3. Create conditional formatting in template Mapping XLS document so folder hierarchy more easy to distinguish e.g.automatic borders or colours for different main folders
    4. Generate empty mapping XLS template to reduce risk of misinterpreting some of the Operation Instructions
    """