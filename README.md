# Translation Files Diff Reporter HK

This is a tool with GUI to compare two source files of translation texts, and see which IDs have difference in source or translation.




## Building executable

pyinstaller --onefile --noconsole --hidden-import=diff_match_patch --upx-dir "c:\Soft\upx-4.0.2-win64" --name hk_diff_reports main.py