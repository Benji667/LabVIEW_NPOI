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

#### Open Document

<section class="body"><iframe src="/LabVIEW NPOI API/Open Document.html" style="border: none" width="960px" height="900px" scrolling="no"></iframe></section>

#### Create New Document

<section class="body"><iframe src="/LabVIEW NPOI API/Create New Document.html" style="border: none" width="960px" height="900px" scrolling="no"></iframe></section>

#### Close Document

<section class="body"><iframe src="/LabVIEW NPOI API/Close Document.html" style="border: none" width="960px" height="750px" scrolling="no"></iframe></section>

## Append Elements

The following VIs allow you to append elements to the document.

#### Append Text

<section class="body"><iframe src="/LabVIEW NPOI API/Append Text.html" style="border: none" width="960px" height="1300px" scrolling="no"></iframe></section>

#### Append Table

* Append Table (string)

<section class="body"><iframe src="/LabVIEW NPOI API/Append Table (string).html" style="border: none" width="960px" height="1500px" scrolling="no"></iframe></section>

* Append Table (double)

<section class="body"><iframe src="/LabVIEW NPOI API/Append Table (double).html" style="border: none" width="960px" height="1500px" scrolling="no"></iframe></section>

*Append Table As Strings (Malleable VI)

<section class="body"><iframe src="/LabVIEW NPOI API/Append Table As Strings.html" style="border: none" width="960px" height="1600px" scrolling="no"></iframe></section>

#### Append Image

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
