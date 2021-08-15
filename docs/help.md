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
			<p><IMG SRC="assets/img/cpath.png" ALT="cpath"></p>
		</TD>
		<TD>
			<P>The <B>File Path</B> specifies the location from where you want to open the **Document** and the name of the **Document**.</P>
		</TD>
	</TR>
	<tr>
		<TD>
			<p>error in (no error)</p>
			<p><IMG SRC="assets/img/cerrcodeclst.png" ALT="cerrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/img/cbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/img/ci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/img/cstr.png" ALT="cstr"></TD>
					<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD>
			<p>Document out</p>
			<P><IMG SRC="assets/img/iLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="iLabVIEW_NPOI_lvlib_Documentlvclass"></P>
		</TD>
		<TD>
			<P>The <B>Document out</B> is a reference to the Document whose appearance, data, and printing you want to control.</P>
		</TD>
	</TR>
			<TR>
		<TD>
			<p>error out</P>
			<P><IMG SRC="assets/img/ierrcodeclst.png" ALT="ierrcodeclst"></P>
		</TD>
		<TD>
			<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
			<P></P>
			<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
			<table class="subtable">
				<TR>
					<TD class="name">status</TD>
					<TD class="terminal"><IMG SRC="assets/img/cbool.png" ALT="cbool"></TD>
					<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
				</TR>
				<TR>
					<TD class="name">code</TD>
					<TD class="terminal"><IMG SRC="assets/img/ci32.png" ALT="ci32"></TD>
					<TD>The <B>code</B> input identifies the error or warning.</TD>
				</TR>
				<TR>
					<TD class="name">source</TD>
					<TD class="terminal"><IMG SRC="assets/img/cstr.png" ALT="cstr"></TD>
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
				<p><IMG SRC="assets/img/cpath.png" ALT="cpath"></p>
			</TD>
			<TD>
				<P>The <B>File Path</B> specifies the location from where you want to open the **Document** and the name of the **Document**.</P>
			</TD>
		</TR>
		<tr>
			<TD>
				<p>error in (no error)</p>
				<p><IMG SRC="assets/img/cerrcodeclst.png" ALT="cerrcodeclst"></P>
			</TD>
			<TD>
				<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
				<P></P>
				<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
				<table class="subtable">
					<TR>
						<TD class="name">status</TD>
						<TD class="terminal"><IMG SRC="assets/img/cbool.png" ALT="cbool"></TD>
						<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
					</TR>
					<TR>
						<TD class="name">code</TD>
						<TD class="terminal"><IMG SRC="assets/img/ci32.png" ALT="ci32"></TD>
						<TD>The <B>code</B> input identifies the error or warning.</TD>
					</TR>
					<TR>
						<TD class="name">source</TD>
						<TD class="terminal"><IMG SRC="assets/img/cstr.png" ALT="cstr"></TD>
						<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
					</TR>
				</table>
			</TD>
		</TR>
		<TR>
			<TD>
				<p>Document out</p>
				<P><IMG SRC="assets/img/iLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="iLabVIEW_NPOI_lvlib_Documentlvclass"></P>
			</TD>
			<TD>
				<P>The <B>Document out</B> is a reference to the Document whose appearance, data, and printing you want to control.</P>
			</TD>
		</TR>
		<TR>
			<TD>
				<p>error out</P>
				<P><IMG SRC="assets/img/ierrcodeclst.png" ALT="ierrcodeclst"></P>
			</TD>
			<TD>
				<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
				<P></P>
				<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
				<table class="subtable">
					<TR>
						<TD class="name">status</TD>
						<TD class="terminal"><IMG SRC="assets/img/cbool.png" ALT="cbool"></TD>
						<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
					</TR>
					<TR>
						<TD class="name">code</TD>
						<TD class="terminal"><IMG SRC="assets/img/ci32.png" ALT="ci32"></TD>
						<TD>The <B>code</B> input identifies the error or warning.</TD>
					</TR>
					<TR>
						<TD class="name">source</TD>
						<TD class="terminal"><IMG SRC="assets/img/cstr.png" ALT="cstr"></TD>
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
				<p><IMG SRC="assets/img/cLabVIEW_NPOI_lvlib_Documentlvclass.png" ALT="cLabVIEW_NPOI_lvlib_Documentlvclass"></p>
			</TD>
			<TD>
				<P><B>Document in</B> is a reference to the Document whose appearance, data, and printing you want to control. Use the &quot;New Document&quot; VI or the &quot;Create Document&quot; VI to generate this LabVIEW class object.</P>
			</TD>
		</TR>
		<tr>
			<TD>
				<p>error in (no error)</p>
				<p><IMG SRC="assets/img/cerrcodeclst.png" ALT="cerrcodeclst"></P>
			</TD>
			<TD>
				<P>The <B>error in</B> cluster can accept error information wired from VIs previously called.  Use this information to decide if any functionality should be bypassed in the event of errors from other VIs.</P>
				<P></P>
				<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
				<table class="subtable">
					<TR>
						<TD class="name">status</TD>
						<TD class="terminal"><IMG SRC="assets/img/cbool.png" ALT="cbool"></TD>
						<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
					</TR>
					<TR>
						<TD class="name">code</TD>
						<TD class="terminal"><IMG SRC="assets/img/ci32.png" ALT="ci32"></TD>
						<TD>The <B>code</B> input identifies the error or warning.</TD>
					</TR>
					<TR>
						<TD class="name">source</TD>
						<TD class="terminal"><IMG SRC="assets/img/cstr.png" ALT="cstr"></TD>
						<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
					</TR>
				</table>
			</TD>
		</TR>
		<TR>
			<TD>
				<p>error out</P>
				<P><IMG SRC="assets/img/ierrcodeclst.png" ALT="ierrcodeclst"></P>
			</TD>
			<TD>
				<P>The <B>error out</B> t cluster passes error or warning information out of a VI to be used by other VIs. </P>
				<P></P>
				<P>The pop-up option <B>Explain Error</B> (or Explain Warning) gives more information about the error displayed. </P>
				<table class="subtable">
					<TR>
						<TD class="name">status</TD>
						<TD class="terminal"><IMG SRC="assets/img/cbool.png" ALT="cbool"></TD>
						<TD>The <B>status</B> boolean is either TRUE (X) for an error, or FALSE (checkmark) for no error or a warning.</TD>
					</TR>
					<TR>
						<TD class="name">code</TD>
						<TD class="terminal"><IMG SRC="assets/img/ci32.png" ALT="ci32"></TD>
						<TD>The <B>code</B> input identifies the error or warning.</TD>
					</TR>
					<TR>
						<TD class="name">source</TD>
						<TD class="terminal"><IMG SRC="assets/img/cstr.png" ALT="cstr"></TD>
						<TD>The <B>source</B> string describes the origin of the error or warning.</TD>
					</TR>
				</table>
			</TD>
		</TR>
	</Table>

## Append Elements

The following VIs allow you to append elements to the document.

+ Append Text

<section class="body"><iframe src="/LabVIEW NPOI API/Append Text.html" style="border: none" width="960px" height="1300px" scrolling="no"></iframe></section>

+ Append Table

- Append Table (string)

<section class="body"><iframe src="/LabVIEW NPOI API/Append Table (string).html" style="border: none" width="960px" height="1500px" scrolling="no"></iframe></section>

-- Append Table (double)

<section class="body"><iframe src="/LabVIEW NPOI API/Append Table (double).html" style="border: none" width="960px" height="1500px" scrolling="no"></iframe></section>

- Append Table As Strings (Malleable VI)

<section class="body"><iframe src="/LabVIEW NPOI API/Append Table As Strings.html" style="border: none" width="960px" height="1600px" scrolling="no"></iframe></section>

+ Append Image

<section class="body"><iframe src="/LabVIEW NPOI API/Append Image.html" style="border: none" width="960px" height="900px" scrolling="no"></iframe></section>

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
