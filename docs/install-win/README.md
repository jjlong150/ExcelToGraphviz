---
prev: /install/
next: /terminology/
---

# Microsoft Windows Installation Instructions

## Install Graphviz

### Step 1 - Download Graphviz Media

Download the installation media from the Windows section of the Graphviz download page at <http://graphviz.org/download/> which appears as:

| ![](../media/ff295e502adb22e77efba6abb86a0507.png) |
| -------------------------------------------------- |

You will be presented with a choice of links. Win32 for the 32-bit installer, and x64 for the 64-bit installer. Choose the link which corresponds with the Windows architecture of your machine.

Download the installer file. At the time of this writing (11-Apr-2023) the files are named:

- **graphviz-8.0.2 (64-bit) EXE installer** [windows_10_cmake_Release_graphviz-install-8.0.2-win64.exe](https://gitlab.com/api/v4/projects/4207231/packages/generic/graphviz-releases/8.0.2/windows_10_cmake_Release_graphviz-install-8.0.2-win64.exe)
- **graphviz-8.0.2 (32-bit) EXE installer** [windows_10_cmake_Release_graphviz-install-8.0.2-win32.exe](https://gitlab.com/api/v4/projects/4207231/packages/generic/graphviz-releases/8.0.2/windows_10_cmake_Release_graphviz-install-8.0.2-win32.exe)

::: tip Download blocked?

If you are using Microsoft Edge, ocassionally the download will be blocked. This is usually due to a new release of Graphviz being published and enough downloads have not occurred to tell Microsoft the download is safe.

Follow the steps on [this page](../blocked/) to learn how to override the block and download the file.

:::

### Step 2 - Launch the installer file.

Click on `Open file` to run the installer.

| ![](../media/7fd0cb6d4e676caea950589d2c1e8cc2.png) |
| -------------------------------------------------- |

A User Access Control warning will now take over the screen and ask

| User Access Control                                                                                                                                                                                                                                                                        |
| :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Do you want to allow this app from an<br>unknown publisher to make changes to your<br>device?<br><br>windows_10_cmake_Release_x64_graphviz-install-8.0.2-win64.exe<br><br>Publisher unknown<br><br>File origin: Hard drive on this computer<br><br>Show more details<br><br>[ Yes ] [ No ] |

Press the `Yes` button, and the Graphviz installer will begin to run.

### Step 3 - Select `Next >`

Select `Next >` on the **"Welcome to Graphviz Setup"** splash page.

![](../media/9adba98571ed4f3aec7cfaaeff63ae56.png)

### Step 4 - Accept the License Agreement

Accept the License Agreement by pressing the `I Agree` button.

![](../media/fd99ccb532cb7aab8242d53a0b0091ed.png)

### Step 5 - Modify PATH

The Graphviz `bin` directory needs to be on your path. You can allow the installer to perform this action (easy way), or you can perform it manually (hard way). We will perform this task manually after the install to ensure you know how to do it should the installer encounter any problems.

Select the radio button `Do not add Graphviz to the system PATH`, then click the `Next >` button.

![](../media/8230e7e73eb6609dce5c24fcc343b6f3.png)

### Step 6 - Specify Installation Folder

Specify the folder where Graphviz should be installed. Select the `Everyone` radio button, and then select the `Next >` button.

![](../media/f50770b78b77a98197663700536363f9.png)

### Step 7 - Press the `Install` button

Choose the Start Menu Folder and press the `Install` button.

![](../media/4dec9767c6961ddf49660e2e69f066d0.png)

### Step 8 - Monitor Progress

The installer will copy the files and make Windows configuration changes. A status bar will indicate how the installation is progressing.

### Step 9 - Click `Finish`

Once the **Completing Graphviz Setup** screen appears press the `Finish` button. The software is installed.

![](../media/2385b02261f96d2a99efe2215cc816e0.png)

## Edit the System PATH Environment Variable

The Relationship Visualizer utilizes the command line programs provided by Graphviz. The Graphviz installer can modify your PATH environment variable, however if your path is too long the installer may fail. It is your responsibility to perform this task manually when following these instructions.

The detailed steps illustrated using a Microsoft Windows 10 operating system are as follows:

### Step 1 - Launch Control Panel

Click the Start button and begin to type "Control Panel". Select the Control Panel icon when it appears.

| ![](../media/267b9de8b8ccf5988f178ad6f72ea8ed.png) |
| -------------------------------------------------- |

### Step 2 - Click on "System and Security"

Click on `System and Security`

| ![](../media/033e8289a0f34c863af9ccd08b6a9bc7.png) |
| -------------------------------------------------- |

### Step 3 - Click on `System`

Click on `System`

| ![](../media/control_panel_system.png) |
| -------------------------------------- |

### Step 4 - Find a Setting

Under “Home”, begin to type “Environment Variables” in the search box “find a setting”.

| ![](../media/find_a_setting.png) |
| -------------------------------- |

Choices beginning with “Env” will begin to appear. Select `Edit the system environment variables`.

| ![](../media/c86551a29e2419d4c4b0be222cb8ed56.png) |
| -------------------------------------------------- |

### Step 5 - Select the "Environment Variables…" button.

Select the `Environment Variables…` button.

| ![](../media/20e421cf752541cf898d91d9f4c2b454.png) |
| -------------------------------------------------- |

### Step 6 - Edit PATH environment variable

“Path” appears in the “User variables” as well as the “System variables”. Highlight the `Path` line in the “Systems variables” list (the bottom list), then select the `Edit…` button.

| ![](../media/8c9a70d533ee0a67b10a0ac6aecb38a2.png) |
| -------------------------------------------------- |

### Step 7 - Press `New`

The **Edit environment variable** dialog appears. Press the `New` button.

| ![](../media/429280385674cae4ec6c208bfd19c053.png) |
| -------------------------------------------------- |

### Step 8 - Press `Browse...`

A new line is added at the bottom of the list. Press the `Browse…` button.

| ![](../media/1b656cc0fc9b7778c5027c2790ba7abc.png) |
| -------------------------------------------------- |

### Step 9 - Navigate to graphviz\bin

The **Browse for Folder** dialog appears. Navigate to the Graphviz bin directory, then press the `OK` button.

| ![](../media/87a0bad2c5d222f055f4ae3a24c610ed.png) |
| -------------------------------------------------- |

If your list is long you may want to use the `Move Up` button to move the directory up in the PATH. Press the `OK` button.

| ![](../media/5c2c576b5b6915ef281da122a420e7fb.png) |
| -------------------------------------------------- |

### Step 10 - Commit the change

You are returned to the **Environment Variables** dialog. Press the `OK` button to commit the environment variable changes and close the dialog.

| ![](../media/b95cc74fdbea30dcc9bd1035bba69593.png) |
| -------------------------------------------------- |

### Step 11 - Restart Windows

(Optional) Restart Microsoft Windows.

Technically this should not be necessary, but if you have already been running Excel there is the possibility that it may be holding an old copy of your environment variables. Restarting Windows will ensure that Excel will reference a current PATH environment variable.

## Perform Graphviz Command Line Configuration and Test

At this point, you have completed the installation steps to install the Graphviz software.

::: warning Do not skip these steps!

The Relationship Visualizer spreadsheet uses the command line programs to generate the graph visualizations. You must manually execute a command line command to configure the Graphviz plugins before you can use Graphviz properly.

:::

Testing the command line programs prior to using the spreadsheet can help ensure that everything is in place correctly so that the spreadsheet can perform properly.

### Step 1 - Open a Command Prompt

Open a Command Prompt window using the `Run as Administrator` option. Click on the Windows Start Menu icon and begin to type Command Prompt. When the Command Prompt App appears choose the `Run as administrator` option.

| ![](../media/f951a48ce6e619cea2dac5d6c14223ce.png) |
| -------------------------------------------------- |

### Step 2 - Run as Administrator

You will get asked for permission to run a command prompt as Administrator.

Press the `Yes` button.

| User Access Control                                                                                                                                                               |
| :-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Do you want to allow this app from an unknown publisher to make changes to your device? <br>Windows Command Processor Verified Publisher: Microsoft Windows<br><br>[ Yes ] [ No ] |

### Step 3 - Confirm Administrator rights

A command prompt window appears. Confirm that the word “Administrator” appears in the Window title.

| ![](../media/d6366c248a6137b8ac6fb4c1e62efc71.png) |
| -------------------------------------------------- |

### Step 4 - Display Graphviz Version

Test that Graphviz is configured properly on the PATH by entering the command:

`dot -V`

noting that the `-V` switch (for version) should be in uppercase, not lowercase. The dot program should respond with the message:

`dot - graphviz version 8.0.2 (20230410.1723)`

in similar fashion to the screen print below:

| ![](../media/5fa025090c64e236250ef11491381a7e.png) |
| -------------------------------------------------- |

If you receive the message:

`'dot' is not recognized as an internal or external command, operable program or batch file.`

It means one of the following things:

- You have specified the path to the Graphviz executables incorrectly and you should repeat the steps above. Things to check are:
  - Did you remember to add the `bin` subdirectory to the Graphviz directory path?
  - Is the directory placed at the end of the `PATH` such that the length of the `PATH` exceeds the Windows length limit? If so, move the Graphviz bin directory closer to the beginning of the list.
- You opened the Command Prompt window prior to updating the PATH statement. This command window is still recognizing the old path. Close the Command Propt window, open a new one, and repeat the `dot -V` command.

### Step 5 - Configure Graphviz Plugins

::: warning
Do not skip this important step!
:::

Configure the plugins by entering the command

`dot -c`

No messages are written when the command executes; the screen will look as follows:

| ![](../media/5064edc9733b86103d396b77bf46229c.png) |
| -------------------------------------------------- |

### Step 6 - View the Plugins List

To see the list of configured plugins type the command

`dot -v`

where the `-v` is lowercase. The screen will appear as follows:

| ![](../media/c78124542c1b1932f338af41a25bad47.png) |
| -------------------------------------------------- |

At this point Graphviz is waiting for more input. Pressing the `Ctrl` + `C` keys will break you from the `dot` program.

### Step 7 - View Command Line Options

To see the list of command line options you can enter the command

`dot -?` or `dot --help`

The screen will appear as follows:

| ![](../media/2f6ee156f09f3ccce9a884967d634429.png) |
| -------------------------------------------------- |

Congratulations! Graphviz is installed properly.

## Install the Relationship Visualizer Excel Spreadsheet Template

### Step 1 - Open Workbook

In the root directory of the Relationship Visualizer distribution there is a macro-enabled Excel spreadsheet named `Relationship Visualizer.xltm`. Double-click the mouse on the file to launch Excel. You will probably get a security warning that the spreadsheet contains macros. You need to enable macro support to use the Relationship Visualizer spreadsheets.

![](../media/d0edc7ed24db40079ea236d5ef0f46f3.png)

### Step 2 - Perform File Save As

Perform a `File` -> `Save As` operation. When it asks you where to store the file, navigate to the directory where you currently have the template file (This PC > Documents > Custom Office Templates).

Next, you will see the file name is `Relationship Visualizer1`. Change it back to `Relationship Visualizer`.

Where it says `Save as type:` select `Excel Macro-Enabled Template` from the dropdown list.

Notice that the save location has changed to your personal `Custom Office Templates` directory.

Select `OK`, and Excel will place a copy of the template file into this directory.

![](../media/23c2e96f733b243b3f5d50c5a8936539.png)

### Step 3 - Close Excel

### Step 4 - Launch Excel

Launch Excel. Excel will offer a selection of built-in spreadsheet templates you can use. Look under the title **FEATURED** for the Relationship Visualizer template.

![](../media/a00c2be365d3e91459f455ba1da09b5f.png)

### Step 5 - Launch Excel Template

The Relationship Visualizer template will be listed, along with a thumbnail image. Click on it.

![](../media/a00c2be365d3e91459f455ba1da09b5f.png)

### Step 6 - Enable Macros to Run

Excel will create a file named **Relationship Visualizer1** with a warning that the spreadsheet contains macros. Click on the `Enable Content` button to permit the macros to run.

![](../media/a6234ff8b7fec03d124a3563e60eeafe.png)

The **SECURITY WARNING** bar should disappear.

![](../media/f60a2ebd0a8ac8b891f2d19fa8854239.png)

### Step 7 - Save File as Macro-Enabled Workbook

Perform a `FILE` -> `Save As` action to save the file as an `Excel Macro-Enabled Workbook`.

Notice that last time we saved it as an `Excel Macro-Enabled Template` but going forward we will create new spreadsheets, populate them with data, and save them as macro-enabled workbooks.

You may save the workbook in the location of your choice, and under the name of your choice. The important thing is saving the file with the `.xlsm` file extension.

![](../media/7a9e903cae1288c61d83db7d9e98e384.png)
