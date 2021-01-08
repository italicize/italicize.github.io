# Make a new Word template

##  Open a document without styles

1. Type **winword /a /w** in the Windows taskbar and press **Enter**. \
<span style='font-size:small;'>&#128712; <span style='color:darkgray;'>The /a switch opens Word without opening your Normal template, which might have custom styles. The /w switch opens a new blank document. For more info see [Command-line switches for Microsoft Office products](https://support.microsoft.com/en-us/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6).</span></span>

## Set the style defaults

1. Click the **Design** menu. 
1. Click **Fonts** and click **Customize Fonts**. 
    1. Select a heading font. \
For a test, select Gill Sans MT.
    1. Select a body font. \
For a test, select Garamond.
    1. Type a name for the theme. \
For a test, type Garamond With Gill Sans.
    1. Click **Save**.
1. Click the **Home** menu.
1. Click the **Styles** button (the small arrow below the Styles gallery). 
1. Click the **Manage Styles** button (the third button at the bottom of the Styles pane).
1. Click the **Set Defaults** tab of the Manage Styles dialog. 
    1. Select the font for tables: \
    &bull; To use the same font as body text, select **+Body**. \
    &bull; To use the same font as headings, select **+Headings**. \
For a test, select +Headings.
    1. Select the size for tables. \
For a test, select 8. 
    1. For the paragraph spacing after, select **0 pt**. 
    1. Type the line spacing for tables. \
For a test, leave the default, Multiple at 1.08.
    1. Click **OK**. 
1. In the Styles pane, right-click **Normal** and select **Modify**. 
    1. Click **Format** and select **Font**. 
    1. Click the **Advanced** tab.
    1. Checkmark **Kerning for fonts** and select **8** points. 
    1. Click **OK**. Click **OK** again.

## Define a new table style

1. Click the **New Style** button (in the lower-left corner of the Styles pane).
1. Type a name. \
For a test, type **Basic Table**. 
1. For the style type, select **Table**. 
    1. Click **Format** and select **Table Properties**.
        1. On the **Table** tab, select an alignment and indent. \
            For a test, select Left and select 0.5" indent (to match the body text indent).
        1. Click **Options**.
        1. Select the default cell margins. \
            For a test, select Top 0.02", Bottom 0.02", Left 0.02", and Right 0.04".
        1. Click **OK**.
        1. Click the **Row** tab. \
            For a test, uncheck "Allow row to break across pages."
        1. Click **OK**.
    1. Click **Format** and select **Borders and Shading**. 
        1. Select borders. \
            For a test, click the top-border and bottom-border buttons.
        1. Click **OK**.
    1. Click **Format** and select **Font**. 
        1. On the **Advanced** tab, checkmark **Kerning for fonts** and select **8** points. 
        1. Click **OK**.
    1. Click **Format** and select **Paragraph**.
        1. On the **Indents and Spacing** tab, select the line spacing. \
            For a test, select Multiple and type 1.08 (instead of 3). 
        1. Click the **Line and Page Breaks** tab. \
            For a test, uncheck Widow/Orphan control.
        1. Click **OK**.
1. For "Apply formatting to," select **Header row**.
    1. Click **Format** and select **Paragraph**.
        1. On the **Line and Page Breaks** tab, checkmark **Keep with next**.
        1. Click **OK**.
    1. Click **Format** and select **Font**.
        1. Click the **Font** tab and select font settings. \
For a test, select Bold.
        1. Click **OK**.
    1. Click **Format** and select **Borders and Shading**. 
        1. Select borders. \
For a test, click the top-border and bottom-border buttons. 
        1. Click **OK**.
    1. Click **Format** and select **Table Properties**.
        1. On the **Row** tab, checkmark **Repeat as header row at the top of each page**.
        1. Click the **Cell** tab.
        1. For the vertical alignment, click **Bottom**.
        1. Click **OK**.
1. Click **OK**.

##  Save the new table style as the default

1. Click the **Insert** menu.
1. Click **Table** and select a **3x3 Table** (or larger). \
A table is inserted. The **Table Tools: Design** menu opens, showing the Table Styles gallery.
1. Right-click the new table style (the first style in the gallery) and select **Set as Default**.
1. Click **OK** to the message "This document only?"
1. Click the **Table Tools: Layout** menu.
1. Click **Delete** and select **Delete Table**.
