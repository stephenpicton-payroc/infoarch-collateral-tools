# Set Up and Use the Pandoc Converter

The Pandoc Converter is a desktop application for converting documents between Word format (.docx) and Markdown format (.md). Before you can use the application, you need to install two programs: Python and Pandoc. You need to install both programs only once.

## Install Python

Python is the programming language that the Pandoc Converter runs on. You don't need to know how to write code to use it.

1. Go to [python.org/downloads](https://www.python.org/downloads).
2. Select **Download Python** to download the latest version.
3. Open the downloaded installer.
4. Select the **Add python.exe to PATH** checkbox.

   **Important:** You must select this checkbox. Without it, the Pandoc Converter can't run when you open it.

5. Select **Install Now**.
6. When the installation is complete, select **Close**.

## Install Pandoc

Pandoc is the conversion tool that the Pandoc Converter uses. You don't need to interact with it directly — the application handles it for you.

1. Go to [pandoc.org/installing.html](https://pandoc.org/installing.html).
2. Under **Windows**, select the **.msi** installer to download it.
3. Open the downloaded installer.
4. Follow the on-screen instructions to complete the installation.

## Set up the Pandoc Converter

To set up the Pandoc Converter, save the application file to your computer and create a shortcut so that you can open it quickly.

1. Copy **PandocConverter.pyw** to a folder on your computer.
2. Right-click **PandocConverter.pyw**, select **Send to**, and then select **Desktop (create shortcut)**.

You can now open the Pandoc Converter by double-clicking the shortcut on your desktop.

## Convert a document

Use the Pandoc Converter to convert a Word document to Markdown, or a Markdown document to Word. The application detects the type of document that you select and converts it to the other format automatically.

1. Double-click the **Pandoc Converter** shortcut on your desktop.
2. Select **Browse**.
3. Find and select the document that you want to convert, and then select **Open**.

   The application shows the conversion direction. For example, if you selected a Word document, it shows **docx → markdown**.

4. Select **Convert**.
5. In the dialog, select the folder where you want to save the converted document.

   **Note:** To save to a new folder, select **New folder**, enter a name for the folder, and then select the folder to open it before you select **Select Folder**.

6. Select **Select Folder**.

The converted document appears in the folder that you selected. The **Log** section at the bottom of the application confirms that the conversion is complete and shows the file path of the converted document.

**Note:** If you converted a Word document, the folder also contains an **images** subfolder. The subfolder contains any images from the original document, and the converted Markdown file links to them automatically.

---


