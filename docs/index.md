---
title: Index
layout: default
---
<div id="top"></div>

[![GitHub all releases][release-shield]][release-url]
[![Wiki][wiki-shield]][wiki-url]
[![Issues][issues-shield]][issues-url]
[![Zero-Clause BSD][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]

# {{ page.title }}

# ![SimpleDocumentCreationExample](https://github.com/Benji667/LabVIEW_NPOI/blob/gp/docs/img/LabVIEW_NPOI_Logo_Small.png?raw=true) [LabVIEW NPOI](https://benji667.github.io/LabVIEW_NPOI/about) 

This project intends to implement a wrapper of the NPOI assembly that allows MS Office files manipulation in LabVIEW (no MS Office installation required).

I'm still exploring the capabilities and possibilities of the NPOI interfaces, trying different approaches. For now this project is more like a PoC. In a second time, I'll try to implement a more reliable , scalable, maintainable architecture. 

Feel free to drop a message in the [Discussions page](https://github.com/Benji667/LabVIEW_NPOI/discussions) and/or contribute.

## Whast is NPOI?

NPOI is a .NET version of POI Java Project. It is an open source .NET library to read and write Microsoft Office file formats formats (*.xls/xlsx, *.doc/docx, *.ppt/pptx). 
You can manually download the repository from [GitHub](https://github.com/nissl-lab/npoi) or install from [NuGet](https://www.nuget.org/packages/NPOI/).

## Installation

* Download the latest version of the LabVIEW NPOI VI Package from the release page.
* Use VIPM to install it on your LabVIEW version (2017+).
* Download and install the latest version of the LUT package available [here](https://github.com/Benji667/LookUp_Table).

## Examples

Below is a simple example to show how to use the LabVIEW NPOI API to interact with Excel or Word document.

![SimpleDocumentCreationExample](https://github.com/Benji667/LabVIEW_NPOI/blob/bcb686f6b338eb219e46d72dd402a0802e551e9f/docs/img/SimpleDocumentCreationExample.png?raw=true)

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

Do not hesitate to go on the [Discussions page](https://github.com/Benji667/LabVIEW_NPOI/discussions).

# Author

BenjaminR

<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[release-shield]: https://img.shields.io/github/v/release/Benji667/LabVIEW_NPOI?color=orange&logo=labview&style=for-the-badge
[release-url]: https://github.com/Benji667/LabVIEW_NPOI/releases/tag/1.1.2
[wiki-shield]: https://img.shields.io/github/discussions/BenjaminRLabVIEWExtensions/Insert-LVClass-Property-Node?style=for-the-badge
[wiki-url]: https://github.com/Benji667/LabVIEW_NPOI/wiki
[issues-shield]: https://img.shields.io/github/issues/BenjaminRLabVIEWExtensions/Insert-LVClass-Property-Node?style=for-the-badge
[issues-url]: https://github.com/Benji667/LabVIEW_NPOI/issues
[license-shield]: https://img.shields.io/badge/LICENSE-Zero--Clause%20BSD-green?style=for-the-badge
[license-url]: https://github.com/Benji667/LabVIEW_NPOI/master/LICENSE
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://www.linkedin.com/in/benjaminrouffet/
