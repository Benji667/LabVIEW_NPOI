---
title: Help
layout: default
---

# {{ page.title }}

# LabVIEW NPOI API

## Overview

Many a time, a software application is required to generate reports in Microsoft Excel file format. Sometimes, an application is even expected to receive Excel files as input data.

### What is LabVIEW NPOI?

LabVIEW NPOI is ab API that allows programmers to create, modify, and display MS Office files using LabVIEW programs. It contains classes and methods to decode the user input data or a file into MS Office documents. 

It attempts to provide an high-level API for an easy and fast intagration by handling the interface to the .NET version of the <A href='http://poi.apache.org/'>POI Java project</A> called <A href='https://github.com/nissl-lab/npoi'>NPOI</A>.

## Prerequisites

<p>
	<ul>
		<li>
			<a href='https://www.ni.com/en-ca/shop/labview.html'>LabVIEW</a> 2017 (Full,Pro or Community) or later.
		</li>
		<li>Operating Systems: <a href='https://www.microsoft.com/en-ca/windows/get-windows-10'>Windows 10</a> (8 but not tested on it).</li>
		<li>.NET Framework 4.0 and above.</li>
	</ul>
</p>

## Installation

<p>
	<ol>
		<li>Download the latest version of the LabVIEW NPOI VI Package from the <A href='https://github.com/Benji667/LabVIEW_NPOI/releases'>release page</A>.</li>
		<li>Use VIPM to install it on your LabVIEW version (2017+).</li>
	</ol>
</p>

## Examples

#### Append String Table To Document.vi
Demonstrates how to add string table to an MS Office document
![Append String Table To Document](https://github.com/Benji667/LabVIEW_NPOI/blob/gp/docs/LabVIEW%20NPOI%20API/img/Append_String_Table_To_Documentp.png?raw=true)

#### Append Image To Document.vi
Demonstrates how to add an image to an MS Office document.
![Append Image To Document](https://github.com/Benji667/LabVIEW_NPOI/blob/gp/docs/LabVIEW%20NPOI%20API/img/Append_Image_To_Documentp.png?raw=true)

#### Append Page To Document Excel.vi
Demonstrates how to add a page to an MS Excel document.
![Append Page To Document Excel](https://github.com/Benji667/LabVIEW_NPOI/blob/gp/docs/LabVIEW%20NPOI%20API/img/Append_New_Page_To_Document_(Excel)p.png?raw=true)

#### Append Text To Document Word.vi
Demonstrates how to add formatted texts to an MS Word document.
![Append Text To Document Word](https://github.com/Benji667/LabVIEW_NPOI/blob/gp/docs/LabVIEW%20NPOI%20API/img/Append_Text_To_Document_(Word)p.png?raw=true)

## Componant API

### Open or Create and Close

The following VIs allow you to Open or Create a (new) document. The call of the Close function is mandatory to close any references that you open or create.

+ Open Document
<P><IMG SRC="assets/img/Open Documentc.png" ALT="LabVIEW NPOI.lvlib:Document.lvclass:Open Document.vi"></P>

Opens the **Document** file located at **File Path** for reading.

Notes:
- If **File Path** is empty (default) or is &lt;Not A Path&gt; or is a directory, error 7 is returned.
- The qualified name of this VI is: NPOI.lvlib:Document.lvclass:Open Document.vi.

<Table style="width:960px">
	<TR style="height:50px">
		<TH style="width:20%"><H3>Terminal</H3></TH>
		<TH><H3>Description</H3></TH>
	</TR>
	<TR>
		<TD>
			<p>File Path</p>
			<p><IMG SRC="assets/imgcpath.png" ALT="cpath"></p>
		</TD>
		<TD>
			<P>The <B>File Path</B> specifies the location from where you want to open the <strong>Document</strong> and the name of the <strong>Document</strong>.</P>
		</TD>
	</TR>
	<tr>
		<TD>
			<p>error in (no error)</p>
			<p><IMG SRC="assets/imgcerrcodeclst.png" ALT="cerrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Document out</p>
			<P><IMG SRC="assets/imgiLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="iLabVIEW_NPOI_lvlib_Documentlvclass"></P>
		</TD>
		<TD>
			<P>The <B>Document out</B> is a reference to the Document whose appearance, data, and printing you want to control.</P>
		</TD>
	</TR>
			<TR>
		<TD>
			<p>error out</P>
			<P><IMG SRC="assets/imgierrcodeclst.png" ALT="ierrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
</Table>

+ Create New Document

<P><IMG SRC="assets/img/Create New Documentc.png" ALT="LabVIEW NPOI.lvlib:Document.lvclass:Create New Document.vi"></P>
			
Creates a new **Document** at the **File Path** location. Use save to persist your changes.

Notes:
- If **File Path** is empty (default) or is &lt;Not A Path&gt; or is a directory, the VI displays a dialog box from which the user can select a file. Error 43 occurs if you cancel the dialog box.
- If the **File Path** already contains a docuement with the same name then that file is overwritten without warning.
- The qualified name of this VI is: NPOI.lvlib:Document.lvclass:Create New Document.vi.
	
<Table style="width:960px">
	<TR style="height:50px">
		<TH style="width:20%"><H3>Terminal</H3></TH>
		<TH><H3>Description</H3></TH>
	</TR>
	<TR>
		<TD>
			<p><B>File Path</B></p>
			<p><IMG SRC="assets/imgcpath.png" ALT="cpath"></p>
		</TD>
		<TD>
			<P>The <B>File Path</B> specifies the location from where you want to open the <strong>Document</strong> and the name of the <strong>Document</strong>.</P>
		</TD>
	</TR>
	<tr>
		<TD>
			<p>error in (no error)</p>
			<p><IMG SRC="assets/imgcerrcodeclst.png" ALT="cerrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Document out</p>
			<P><IMG SRC="assets/imgiLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="iLabVIEW_NPOI_lvlib_Documentlvclass"></P>
		</TD>
		<TD>
			<P>The <B>Document out</B> is a reference to the Document whose appearance, data, and printing you want to control.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>error out</P>
			<P><IMG SRC="assets/imgierrcodeclst.png" ALT="ierrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
</Table>

+ Close Document
<P><IMG SRC="assets/img/Close Documentc.png" ALT="LabVIEW NPOI.lvlib:Document.lvclass:Close Document.vi"></P>

Closes any references that you open or create for the **Document**.

Notes:
- The qualified name of this VI is: NPOI.lvlib:Document.lvclass:Close Document.vi.

<Table style="width:960px">
	<TR style="height:50px">
		<TH style="width:20%"><H3>Terminal</H3></TH>
		<TH><H3>Description</H3></TH>
	</TR>
	<TR>
		<TD>
			<p><B>Document in</B></p>
			<p><IMG SRC="assets/imgcLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="cLabVIEW_NPOI_lvlib_Documentlvclass"></p>
		</TD>
		<TD>
			<P><B>Document in</B> is a reference to the Document whose appearance, data, and printing you want to control. Use the &quot;New Document&quot; VI or the &quot;Create Document&quot; VI to generate this LabVIEW class object.</P>
		</TD>
	</TR>
	<tr>
		<TD>
			<p>error in (no error)</p>
			<p><IMG SRC="assets/imgcerrcodeclst.png" ALT="cerrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>error out</P>
			<P><IMG SRC="assets/imgierrcodeclst.png" ALT="ierrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
</Table>

## Append Elements

The following VIs allow you to append elements to the document.

+ Append Text
<p>
	<DL>
		<DT>
			<P><IMG SRC="assets/img/Append Textc.png" ALT="LabVIEW NPOI.lvlib:Document.lvclass:Append Text.vi"></P>
			<P></P>
			<P>Appends <strong>Text</strong> to the <strong>Document</strong>. </P>
		</DT>
		<DT>
		<P><i>Notes</i> :</P>
		</DT>
			<DD>
				<p>- The qualified name of this VI is: NPOI.lvlib:Document.lvclass:Append Text.vi.</p>
			</DD>
	</DL>
</p>
<Table style="width:960px">
	<TR style="height:50px">
		<TH style="width:20%"><H3>Terminal</H3></TH>
		<TH><H3>Description</H3></TH>
	</TR>
	<TR>
		<TD>
			<p><B>Document in</B></p>
			<p><IMG SRC="assets/imgcLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="cLabVIEW_NPOI_lvlib_Documentlvclass"></p>
		</TD>
		<TD>
			<P><B>Document in</B> is a reference to the Document whose appearance, data, and printing you want to control. Use the &quot;New Document&quot; VI or the &quot;Create Document&quot; VI to generate this LabVIEW class object.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p><B>Text input</B></p>
			<p><IMG SRC="assets/imgcstr.png" ALT="cstr"></p>
		</TD>
		<TD>
			<P><B>Text</B> is the information you want to include in the Document.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Style</p>
			<p><IMG SRC="assets/imgccclst.png" ALT="ccclst"></p>
		</TD>
		<TD>
			<P><strong>Style</strong> indicates how the text appears in the Document.</P>
			<table class="subtable">
				<TR>
					<TD class="name">Name</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD colspan="2">The <B>Name</B> is the style name.</TD>
				</TR>
				<TR>
					<TD class="name">Font</TD>
					<TD class="terminal"><IMG SRC="assets/imgccclst.png" ALT="ccclst"></TD>
					<TD colspan="2">The <B>Font</B> indicates the font settings used for the Paragraph.</TD>
				</TR>
				<TR>
					<TD> </TD>
					<TD colspan="2">
						<BR>
						The <B>Font</B> cluster is composed of:
					</TD>
				</TR>
				<TR>
					<TD> </TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Name</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>Name</B> indicates the name of the font used, such as Times New Roman.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Size</TD>
					<td class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>Size</B> indicates the size of the font used.</TD>	
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Bold</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Bold</B> indicates whether the text is in bold.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Italic</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Italic</B> indicates whether the text is in italics.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Underline</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Underline</B> indicates whether the text is underlined.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Strike</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Strike</B> indicates whether the text is struck through.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Color</TD>
					<td class="terminal"><IMG SRC="assets/imgcu32.png" ALT="cu32"></TD>
					<TD>The <B>Color</B> indicates the color of the text.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Alignment</TD>
					<td class="terminal"><IMG SRC="assets/imgcu32.png" ALT="cu32"></TD>
					<TD>The <B>Alignment</B> indicates the text alignment.</TD>
				</TR>
				<TR>
					<TD><BR></TD>
				</TR>
			</Table>
		</TD>
	</TR>

	</tr>
	<tr>
		<TD>
			<p>error in (no error)</p>
			<p><IMG SRC="assets/imgcerrcodeclst.png" ALT="cerrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Document out</p>
			<P><IMG SRC="assets/imgiLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="iLabVIEW_NPOI_lvlib_Documentlvclass"></P>
		</TD>
		<TD>
			<P>The <B>Document out</B> is a reference to the Document whose appearance, data, and printing you want to control.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>error out</P>
			<P><IMG SRC="assets/imgierrcodeclst.png" ALT="ierrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
</Table>

+ Append Table

- Append Table (string)

<DL>
	<DT>
		<P><IMG SRC="assets/imgAppend Table (string)c.png" ALT="LabVIEW NPOI.lvlib:Document.lvclass:Append Table (string).vi"></P>
		<P></P>
		<P>Appends the wired <strong>Table</strong> as 2D array of strings to the <strong>Document</strong> as a table. Wire data to the <B>Table</B> input to determine the polymorphic instance to use or manually select the instance.</P>
	</DT>
	<DT>
	<P><i>Notes</i> :</P>
	</DT>
		<DD>
			<p>- The qualified name of this VI is: NPOI.lvlib:Document.lvclass:Append Table (string).vi.</p>
		</DD>
</DL>
<Table style="width:960px">
	<TR style="height:50px">
		<TH style="width:20%"><H3>Terminal</H3></TH>
		<TH><H3>Description</H3></TH>
	</TR>
	<TR>
		<TD>
			<p><B>Document in</B></p>
			<p><IMG SRC="assets/imgcLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="cLabVIEW_NPOI_lvlib_Documentlvclass"></p>
		</TD>
		<TD>
			<P><B>Document in</B> is a reference to the Document whose appearance, data, and printing you want to control. Use the &quot;New Document&quot; VI or the &quot;Create Document&quot; VI to generate this LabVIEW class object.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p><B>Table</B></p>
			<p><IMG SRC="assets/imgc2dstr.png" ALT="cstr"></p>
		</TD>
		<TD>
			<P><B>Table</B> contains the data of the table inserted into the <B>Docuement</B>.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Column Headers</p>
			<p><IMG SRC="assets/imgc1dstr.png" ALT="c1dstr"></p>
		</TD>
		<TD>
			<P><B>Column Headers</B> determines how each column is labeled in the table. </P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Row Headers</p>
			<p><IMG SRC="assets/imgc1dstr.png" ALT="c1dstr"></p>
		</TD>
		<TD>
			<P><B>Row Headers</B> determines how each row is labeled in the table. </P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Style</p>
			<p><IMG SRC="assets/imgccclst.png" ALT="ccclst"></p>
		</TD>
		<TD>
			<P><strong>Style</strong> indicates how the text appears in the Document.</P>
			<table class="subtable">
				<TR>
					<TD class="name">Name</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD colspan="2">The <B>Name</B> is the style name.</TD>
				</TR>
				<TR>
					<TD class="name">Font</TD>
					<TD class="terminal"><IMG SRC="assets/imgccclst.png" ALT="ccclst"></TD>
					<TD colspan="2">The <B>Font</B> indicates the font settings used for the Paragraph.</TD>
				</TR>
				<TR>
					<TD> </TD>
					<TD colspan="2">
						<BR>
						The <B>Font</B> cluster is composed of:
					</TD>
				</TR>
				<TR>
					<TD> </TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Name</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>Name</B> indicates the name of the font used, such as Times New Roman.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Size</TD>
					<td class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>Size</B> indicates the size of the font used.</TD>	
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Bold</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Bold</B> indicates whether the text is in bold.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Italic</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Italic</B> indicates whether the text is in italics.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Underline</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Underline</B> indicates whether the text is underlined.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Strike</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Strike</B> indicates whether the text is struck through.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Color</TD>
					<td class="terminal"><IMG SRC="assets/imgcu32.png" ALT="cu32"></TD>
					<TD>The <B>Color</B> indicates the color of the text.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Alignment</TD>
					<td class="terminal"><IMG SRC="assets/imgcu32.png" ALT="cu32"></TD>
					<TD>The <B>Alignment</B> indicates the text alignment.</TD>
				</TR>
				<TR>
					<TD><BR></TD>
				</TR>
			</Table>
		</TD>
	</TR>

	</tr>
	<tr>
		<TD>
			<p>error in (no error)</p>
			<p><IMG SRC="assets/imgcerrcodeclst.png" ALT="cerrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Document out</p>
			<P><IMG SRC="assets/imgiLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="iLabVIEW_NPOI_lvlib_Documentlvclass"></P>
		</TD>
		<TD>
			<P>The <B>Document out</B> is a reference to the Document whose appearance, data, and printing you want to control.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>error out</P>
			<P><IMG SRC="assets/imgierrcodeclst.png" ALT="ierrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
</Table>

-- Append Table (double)

<DL>
	<DT>
		<A NAME="LabVIEW NPOI.lvlib:Document.lvclass:Append Table _double_.vi"></A>
		<H2>Append Table (double).vi</H2>
		<P><IMG SRC="assets/imgAppend Table (double)c.png" ALT="LabVIEW NPOI.lvlib:Document.lvclass:Append Table (double).vi"></P>
		<P></P>
		<P>Appends the wired <strong>Table</strong> as 2D array of doubles to the <strong>Document</strong> as a table. Wire data to the <B>Table</B> input to determine the polymorphic instance to use or manually select the instance.</P>
	</DT>
	<DT>
	<P><i>Notes</i> :</P>
	</DT>
		<DD>
			<p>- The qualified name of this VI is: NPOI.lvlib:Document.lvclass:Append Table (double).vi.</p>
		</DD>
</DL>
<Table style="width:960px">
	<TR style="height:50px">
		<TH style="width:20%"><H3>Terminal</H3></TH>
		<TH><H3>Description</H3></TH>
	</TR>
	<TR>
		<TD>
			<p><B>Document in</B></p>
			<p><IMG SRC="assets/imgcLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="cLabVIEW_NPOI_lvlib_Documentlvclass"></p>
		</TD>
		<TD>
			<P><B>Document in</B> is a reference to the Document whose appearance, data, and printing you want to control. Use the &quot;New Document&quot; VI or the &quot;Create Document&quot; VI to generate this LabVIEW class object.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p><B>Table</B></p>
			<p><IMG SRC="assets/imgcdbl.png" ALT="cdbl"></p>
		</TD>
		<TD>
			<P><B>Table</B> contains the data of the table inserted into the <B>Docuement</B>.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Column Headers</p>
			<p><IMG SRC="assets/imgc1dstr.png" ALT="c1dstr"></p>
		</TD>
		<TD>
			<P><B>Column Headers</B> determines how each column is labeled in the table. </P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Row Headers</p>
			<p><IMG SRC="assets/imgc1dstr.png" ALT="c1dstr"></p>
		</TD>
		<TD>
			<P><B>Row Headers</B> determines how each row is labeled in the table. </P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Style</p>
			<p><IMG SRC="assets/imgccclst.png" ALT="ccclst"></p>
		</TD>
		<TD>
			<P><strong>Style</strong> indicates how the text appears in the Document.</P>
			<table class="subtable">
				<TR>
					<TD class="name">Name</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD colspan="2">The <B>Name</B> is the style name.</TD>
				</TR>
				<TR>
					<TD class="name">Font</TD>
					<TD class="terminal"><IMG SRC="assets/imgccclst.png" ALT="ccclst"></TD>
					<TD colspan="2">The <B>Font</B> indicates the font settings used for the Paragraph.</TD>
				</TR>
				<TR>
					<TD> </TD>
					<TD colspan="2">
						<BR>
						The <B>Font</B> cluster is composed of:
					</TD>
				</TR>
				<TR>
					<TD> </TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Name</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>Name</B> indicates the name of the font used, such as Times New Roman.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Size</TD>
					<td class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>Size</B> indicates the size of the font used.</TD>	
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Bold</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Bold</B> indicates whether the text is in bold.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Italic</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Italic</B> indicates whether the text is in italics.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Underline</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Underline</B> indicates whether the text is underlined.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Strike</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Strike</B> indicates whether the text is struck through.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Color</TD>
					<td class="terminal"><IMG SRC="assets/imgcu32.png" ALT="cu32"></TD>
					<TD>The <B>Color</B> indicates the color of the text.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Alignment</TD>
					<td class="terminal"><IMG SRC="assets/imgcu32.png" ALT="cu32"></TD>
					<TD>The <B>Alignment</B> indicates the text alignment.</TD>
				</TR>
				<TR>
					<TD><BR></TD>
				</TR>
			</Table>
		</TD>
	</TR>

	</tr>
	<tr>
		<TD>
			<p>error in (no error)</p>
			<p><IMG SRC="assets/imgcerrcodeclst.png" ALT="cerrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Document out</p>
			<P><IMG SRC="assets/imgiLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="iLabVIEW_NPOI_lvlib_Documentlvclass"></P>
		</TD>
		<TD>
			<P>The <B>Document out</B> is a reference to the Document whose appearance, data, and printing you want to control.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>error out</P>
			<P><IMG SRC="assets/imgierrcodeclst.png" ALT="ierrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
</Table>

- Append Table As Strings (Malleable VI)

<DL>
	<DT>
		<A NAME="LabVIEW NPOI.lvlib:Document.lvclass:Append Table As Strings.vim"></A>
		<H2>Append Table As Strings.vim</H2>
		<P><IMG SRC="assets/imgAppend Table As Stringsc.png" ALT="Append Table As Strings.vim"></P>
		<P></P>
		<P>Appends the wired <B>Table</B> as strings to the <B>Document</B>.</P>
	</DT>
	<DT>
	<P><i>Notes</i> :</P>
	</DT>
		<DD>
			<P>- As a malleable VI (.vim) this VI is inlined into its calling VI and can adapt each terminal to its corresponding input data type.</P>
		</DD>
		<DD>
			<P>- The following data types will be formatted to strings: 2D or 1D array of strings, all numerics (except Fixed-Point), timestamp, and Boolean. All other data types will cause a broken wire and broken run arrow.</P>
		</DD>
		<DD>
			<p>- The qualified name of this VI is: NPOI.lvlib:Document.lvclass:Append Table As Strings.vim.</p>
		</DD>
</DL>
<Table style="width:960px">
	<TR style="height:50px">
		<TH style="width:20%"><H3>Terminal</H3></TH>
		<TH><H3>Description</H3></TH>
	</TR>
	<TR>
		<TD>
			<p><B>Document in</B></p>
			<p><IMG SRC="assets/imgcLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="cLabVIEW_NPOI_lvlib_Documentlvclass"></p>
		</TD>
		<TD>
			<P><B>Document in</B> is a reference to the Document whose appearance, data, and printing you want to control. Use the &quot;New Document&quot; VI or the &quot;Create Document&quot; VI to generate this LabVIEW class object.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p><B>Table</B></p>
			<p><IMG SRC="assets/imgc2dstr.png" ALT="cstr"></p>
		</TD>
		<TD>
			<P><B>Table</B> contains the data of the table inserted into the <B>Docuement</B>.</P>
			<P>This input accepts a 2D or 1D array of strings, all numerics (except Fixed-Point), timestamp, and Boolean. All other data types will cause a broken wire and broken run arrow.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Column Headers</p>
			<p><IMG SRC="assets/imgc1dstr.png" ALT="c1dstr"></p>
		</TD>
		<TD>
			<P><B>Column Headers</B> determines how each column is labeled in the table. </P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Row Headers</p>
			<p><IMG SRC="assets/imgc1dstr.png" ALT="c1dstr"></p>
		</TD>
		<TD>
			<P><B>Row Headers</B> determines how each row is labeled in the table. </P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Style</p>
			<p><IMG SRC="assets/imgccclst.png" ALT="ccclst"></p>
		</TD>
		<TD>
			<P><strong>Style</strong> indicates how the text appears in the Document.</P>
			<table class="subtable">
				<TR>
					<TD class="name">Name</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD colspan="2">The <B>Name</B> is the style name.</TD>
				</TR>
				<TR>
					<TD class="name">Font</TD>
					<TD class="terminal"><IMG SRC="assets/imgccclst.png" ALT="ccclst"></TD>
					<TD colspan="2">The <B>Font</B> indicates the font settings used for the Paragraph.</TD>
				</TR>
				<TR>
					<TD> </TD>
					<TD colspan="2">
						<BR>
						The <B>Font</B> cluster is composed of:
					</TD>
				</TR>
				<TR>
					<TD> </TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Name</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>Name</B> indicates the name of the font used, such as Times New Roman.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Size</TD>
					<td class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>Size</B> indicates the size of the font used.</TD>	
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Bold</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Bold</B> indicates whether the text is in bold.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Italic</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Italic</B> indicates whether the text is in italics.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Underline</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Underline</B> indicates whether the text is underlined.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Strike</TD>
					<td class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>Strike</B> indicates whether the text is struck through.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Color</TD>
					<td class="terminal"><IMG SRC="assets/imgcu32.png" ALT="cu32"></TD>
					<TD>The <B>Color</B> indicates the color of the text.</TD>
				</TR>
				<TR>
					<TD></TD>
					<TD class="name">Alignment</TD>
					<td class="terminal"><IMG SRC="assets/imgcu32.png" ALT="cu32"></TD>
					<TD>The <B>Alignment</B> indicates the text alignment.</TD>
				</TR>
				<TR>
					<TD><BR></TD>
				</TR>
			</Table>
		</TD>
	</TR>

	</tr>
	<tr>
		<TD>
			<p>error in (no error)</p>
			<p><IMG SRC="assets/imgcerrcodeclst.png" ALT="cerrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Document out</p>
			<P><IMG SRC="assets/imgiLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="iLabVIEW_NPOI_lvlib_Documentlvclass"></P>
		</TD>
		<TD>
			<P>The <B>Document out</B> is a reference to the Document whose appearance, data, and printing you want to control.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>error out</P>
			<P><IMG SRC="assets/imgierrcodeclst.png" ALT="ierrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
</Table>

+ Append Image

<DL>
<DT>
	<A NAME="LabVIEW NPOI.lvlib:Document.lvclass:Append Image.vi"></A>
	<H2>Append Image.vi</H2>
	<P><IMG SRC="assets/imgAppend Imagec.png" ALT="LabVIEW NPOI.lvlib:Document.lvclass:Append Image.vi"></P>
	<P></P>
	<P>Appends the image located at <strong>Image File Path</strong> to the <strong>Document</strong>. Only PNG and JPG format are supported.</P>
</DT>
<DT>
	<P><i>Notes</i> :</P>
</DT>
	<DD>
		<p>- The qualified name of this VI is: NPOI.lvlib:Document.lvclass:Append Image.vi</p>
	</DD>
</DL>
<Table style="width:960px">
	<TR style="height:50px">
		<TH style="width:20%"><H3>Terminal</H3></TH>
		<TH><H3>Description</H3></TH>
	</TR>
	<TR>
		<TD>
			<p><B>Document in</B></p>
			<p><IMG SRC="assets/imgcLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="cLabVIEW_NPOI_lvlib_Documentlvclass"></p>
		</TD>
		<TD>
			<P><B>Document in</B> is a reference to the Document whose appearance, data, and printing you want to control. Use the &quot;New Document&quot; VI or the &quot;Create Document&quot; VI to generate this LabVIEW class object.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p><B>Image File Path</B></p>
			<p><IMG SRC="assets/imgcpath.png" ALT="cpath"></p>
		</TD>
		<TD>
			<P><B>Image File Path</B> designates the path to the linked image. If you move the image, you must update the path.</P>
		</TD>
	</TR>
	<tr>
		<TD>
			<p>error in (no error)</p>
			<p><IMG SRC="assets/imgcerrcodeclst.png" ALT="cerrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Document out</p>
			<P><IMG SRC="assets/imgiLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="iLabVIEW_NPOI_lvlib_Documentlvclass"></P>
		</TD>
		<TD>
			<P>The <B>Document out</B> is a reference to the Document whose appearance, data, and printing you want to control.</P>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>error out</P>
			<P><IMG SRC="assets/imgierrcodeclst.png" ALT="ierrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/imgcbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/imgci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/imgcstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
</Table>

## Retrieve Elements

The following VIs allow you to retrieve elements from the document.

#### Read Text

<section class="body"><iframe src="/LabVIEW NPOI API/Read Text.html" style="border: none" width="960px" height="900px" scrolling="no"></iframe></section>

#### Read Table

* Read Table As String

<section class="body"><iframe src="/LabVIEW NPOI API/Read Table As String.html" style="border: none" width="960px" height="950px" scrolling="no"></iframe></section>

* Read Table (string)

<section class="body"><iframe src="/LabVIEW NPOI API/Retrieve Table (string).html" style="border: none" width="960px" height="900px" scrolling="no"></iframe></section>

* Read Table (double)

<section class="body"><iframe src="/LabVIEW NPOI API/Retrieve Table (double).html" style="border: none" width="960px" height="900px" scrolling="no"></iframe></section>

## Properties

#### Document Properties

<section class="body"><iframe src="/LabVIEW NPOI API/Document Properties.html" style="border: none" width="960px" height="600px" scrolling="no"></iframe></section>

#### Document Property Node

<section class="body"><iframe src="/LabVIEW NPOI API/Document Property Node.html" style="border: none" width="960px" height="1000px" scrolling="no"></iframe></section>

## Excel Specific

The following VIs allow you to incorporate Microsoft Excel features into the document.

#### Add New Page

<section class="body"><iframe src="/LabVIEW NPOI API/Append New Page.html" style="border: none" width="960px" height="1100px" scrolling="no"></iframe></section>

#### Remove Page

<section class="body"><iframe src="/LabVIEW NPOI API/Remove Page.html" style="border: none" width="960px" height="1000px" scrolling="no"></iframe></section>

## Save And Print

The following VIs allow you to save or print the document.

#### Save Document

<section class="body"><iframe src="/LabVIEW NPOI API/Save Document.html" style="border: none" width="960px" height="850px" scrolling="no"></iframe></section>

#### Save As Document

<section class="body"><iframe src="/LabVIEW NPOI API/SaveAs Document.html" style="border: none" width="960px" height="900px" scrolling="no"></iframe></section>

#### Convert To HTML

<section class="body"><iframe src="/LabVIEW NPOI API/Convert To HTML.html" style="border: none" width="960px" height="1000px" scrolling="no"></iframe></section>

#### Print Document

<section class="body"><iframe src="/LabVIEW NPOI API/Print Document.html" style="border: none" width="960px" height="1000px" scrolling="no"></iframe></section>

## Legal

<section class="body"><iframe src="/LabVIEW NPOI API/Legal.html" style="border: none" width="960px" height="200px" scrolling="no"></iframe></section>

<!--
You can use HTML elements in Markdown, such as the comment element, and they won't
be affected by a markdown parser. However, if you create an HTML element in your
markdown file, you cannot use markdown syntax within that element's contents.
-->
 
