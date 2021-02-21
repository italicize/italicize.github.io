# Style specifications

The macro reads specific words. The specifications can appear in any order, except for Styles gallery and _ list.

#### `Margins, . . .` defines the margins.

For example, the Microsoft default margins are **Margins, 1" left, 1" right, 1" top, 1" bottom, 0.5" header, 0.5" footer.**

Specification|Examples|Notes
:---|:---|:---
\_&nbsp;left,<br>\_&nbsp;right,<br>\_&nbsp;top,<br>\_&nbsp;bottom,<br>\_&nbsp;header,<br>\_&nbsp;footer|0.6&quot;&nbsp;left,<br>1&nbsp;inch&nbsp;top|The number at the beginning is read to set a margin in inches (whether inch or in. or &Prime; is included or not). Currently the macro doesn't look for a unit, such as cm or pt.
mirror&nbsp;margins,<br>no&nbsp;mirror&nbsp;margins||Swaps the left and right margins on odd and even pages<br>(or doesn't swap if "no" at the beginning). 

#### `Defaults for all defined styles, . . .` defines all styles.<br>

The available specifications are the same as for style definitions (see next).<br>
Style definitions don't need to repeat any defaults, which allows style definitions to be shorter.<br>
Style definitions can include on the specifications that supersede the defaults.

#### `_ style, . . .` defines a style.<br>
For example, a Microsoft default style is **Body Text style, body font, 11 pt, 1.08 line spacing, 6 pt after, widow/orphan control, no kerning, based on Normal, followed by Body Text.**

#### `Styles gallery, . . .` defines the Styles gallery in the Home menu.

For example, the Microsoft default Styles gallery is **Styles gallery, Normal, No Spacing, Heading 1, Heading 2, Heading 3, Heading 4, Heading 5, Heading 6, Heading 7, Heading 8, Heading 9, Title, Subtitle, Subtle Emphasis, Emphasis, Intense Emphasis, Strong, Quote, Intense Quote, Subtle Reference, Intense Reference, Book Title, List Paragraph, Caption, TOC Heading.**

#### `_ list, . . .` defines a list name and the styles in a list.

The beginning is read as list name. Several variations are accepted, such as "_ multilevel list styles."<br>
The macro adds a new list, if it doesn't exist. The specifications are read as style names, in order, up to nine.<br>
For example, **ListBullets list, List Bullet, List Bullet 2, List Bullet 3, List Bullet 4, List Bullet 5.**

#### `_ bullet defaults, . . .` or `_ number defaults, . . .` defines all bullets or numbers in a list.

The beginning is read as list name. The specifications . . . . . . . <br>
For example, **ListBullets bullet defaults, Body bullet font, tab after bullet.**

#### Specifications for paragraph styles and character styles

Specification|Examples|Notes
:---|:---|:---
based&nbsp;on&nbsp;\_ |based&nbsp;on&nbsp;no&nbsp;style,<br>based&nbsp;on&nbsp;Heading&nbsp;1|The end is read as a style name, the style with specifications to copy. The Microsoft defaults are "based on Normal" for paragraph styles and "based on Default Paragraph Font" for font styles.
\_&nbsp;font|Palatino&nbsp;Linotype&nbsp;font,<br>body&nbsp;font,<br>headings&nbsp;font|The beginning is read as a font name. "Body font" and "headings font" use the defaults (defined through the Design menu).
\_&nbsp;size|11 pt size|The number at the beginning is read as the font size in points (whether pt or point is included or not).
bold,<br>not bold||
italic,<br>not italic||
small caps||
all caps||
\_&nbsp;color|#808080&nbsp;color,<br>black&nbsp;color,<br>automatic&nbsp;color|The number at the beginning is read as a hex value. Words after the number are ignored, such as "#808080 gray color." Currently the macro reads the words "automatic" and "black" but not other colors.
\_&nbsp;character&nbsp;spacing|normal&nbsp;character&nbsp;spacing,<br>0.5 pt&nbsp;character&nbsp;spacing|The number at the beginning is read as points. "Normal" is read as 0 pt, meaning no extra space or reduced space between characters.
kerning,<br>no&nbsp;kerning||
        
#### Specifications for paragraph styles only

Specification|Examples|Notes
:---|:---|:---
followed&nbsp;by&nbsp;\_ |followed&nbsp;by&nbsp;Body&nbsp;Text|The end is read as a style name, the style for the next paragraph after pressing Enter.
space&nbsp;between,<br>no&nbsp;space&nbsp;between||Several variations are accepted, such as "add space between paragraphs of the same style" and "don't add space between paragraphs." No space between is Microsoft's default setting for bullet lists.
\_&nbsp;left&nbsp;indent,<br>\_&nbsp;right&nbsp;indent|0.5&quot;&nbsp;left&nbsp;indent,<br>-0.05&nbsp;right&nbsp;indent|The number at the beginning is read as an indent or outdent from the margin in inches (whether inch or in. or &Prime; is included or not). Currently the macro doesn't look for a unit, such as cm or pt.
\_&nbsp;before,<br>\_&nbsp;after|6 pt after|The number at the beginning is read as points for the space before or after a paragraph. Both before and after can be defined, but defining the space after is enough.
\_&nbsp;line&nbsp;spacing|1.08&nbsp;line&nbsp;spacing,<br>12&nbsp;pt&nbsp;line&nbsp;spacing,<br>at&nbsp;least&nbsp;10.5&nbsp;pt&nbsp;line&nbsp;spacing|The number at the beginning is read as the number of lines or, if "pt" appears, as the number of points of line spacing. The words "exact" or "at least" can appear before the number of points.
left&nbsp;aligned,<br>right&nbsp;aligned,<br>centered,<br>justified||Several variations are accepted, such as "align left" and "center align" and "justify."
widow/orphan&nbsp;control,<br>no&nbsp;widow/orphan&nbsp;control||Several variations are accepted, such as "widow control" and "no widow or orphan control."
keep&nbsp;with&nbsp;next,<br>don't&nbsp;keep&nbsp;with&nbsp;next||Several variations are accepted, such as "keep the paragraph with the next paragraph" and "allow a page break after the paragraph."
keep&nbsp;lines&nbsp;together,<br>don't&nbsp;keep&nbsp;lines&nbsp;together||Several variations are accepted, such as "keep the paragraph lines together" and "allow a page break within the paragraph."
page&nbsp;break&nbsp;before,<br>no&nbsp;page&nbsp;break&nbsp;before||Several variations are accepted, such as "page break above the paragraph" and "don't require a page break before the paragraph."
\_&nbsp;top&nbsp;border,<br>\_&nbsp;bottom&nbsp;border,<br>\_&nbsp;left&nbsp;border,<br>\_&nbsp;right&nbsp;border|1 pt top border|The number at the beginning is read as the border line width in points.
\_&nbsp;left&nbsp;tab,<br>\_&nbsp;center&nbsp;tab,<br>\_&nbsp;right&nbsp;tab,<br>center&nbsp;tab,<br>right&nbsp;tab,<br>no&nbsp;tabs|1&quot;&nbsp;left&nbsp;tab|The number at the beginning is read as a tab position in inches (whether inch or in. or &Prime; is included or not). If no number appears, then "center tab" is centered between the page margins and "right tab" is at the right margin.

#### Specifications for paragraph styles in lists

Specification|Examples|Notes
:---|:---|:---
\_&nbsp;bullet,<br>no&nbsp;bullet|&bull;&nbsp;bullet|The character at the beginning of the line is read as a bullet.
"\_"&nbsp;\_&nbsp;number,<br>"\_"&nbsp;\_&nbsp;letter|&ldquo;%1&rdquo;&nbsp;Roman&nbsp;number,<br>&ldquo;%1.%2.%3.&rdquo;&nbsp;number,<br>&ldquo;(%2)&rdquo;&nbsp;lowercase&nbsp;letter|The text inside quotation marks is read as a number code. %1 means the first-level numbering of a multilevel list, %2 means the second level, and so on. The descriptive word (Roman, lowercase, uppercase) applies to the last level only. Arabic numbers are the default. The effect is the same whether the word "number" or "letter" is used.
tab&nbsp;after&nbsp;bullet,<br>space&nbsp;after&nbsp;bullet,<br>nothing&nbsp;after&nbsp;bullet|tab&nbsp;after&nbsp;number,<br>space&nbsp;after&nbsp;number,<br>nothing&nbsp;after&nbsp;number|The effect is the same whether the word "bullet" or "number" is used. Several variations are accepted, such as "follows" or "following" instead of "after."
\_&nbsp;bullet&nbsp;font,<br>\_&nbsp;number&nbsp;font|Segoe&nbsp;UI&nbsp;bullet&nbsp;font,<br>body&nbsp;bullet&nbsp;font,<br>headings&nbsp;number&nbsp;font|The beginning is read as a font name for the bullet or number. "Body font" and "headings font" use the defaults (defined through the Design menu). 
bold&nbsp;numbers,<br>italic&nbsp;numbers,<br>bold&nbsp;and&nbsp;italic&nbsp;numbers||The effect is the same whether the word "bullet" or "number" is used.
\_&nbsp;bullet&nbsp;color,<br>\_&nbsp;number&nbsp;color|#808080&nbsp;bullet&nbsp;color,<br>black&nbsp;number&nbsp;color,<br>automatic&nbsp;number&nbsp;color|The number at the beginning is read as a hex value. Words after the number are ignored, such as "#808080 gray color." Currently the macro reads the words "automatic" and "black" but not other colors.
\_&nbsp;bullet&nbsp;indent,<br>\_&nbsp;number&nbsp;indent,<br>\_&nbsp;text&nbsp;indent||The number at the beginning is read as an indent or outdent from the margin in inches (whether inch or in. or &Prime; is included or not). Currently the macro doesn't look for a unit, such as cm or pt.

