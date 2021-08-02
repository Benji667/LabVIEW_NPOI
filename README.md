# NPOI assembly wrapper in LabVIEW.

This project intends to implement a wrapper of the NPOI assembly that allows MS Office files manipulation in LabVIEW (no MS Office installation required).
I'm still exploring NPOI interface for capabilities and possibilities; trying different approaches, more like a PoC. In a second time, I'll try to implement a reliable , scalable, maintainable architecture. 
Feel free to let me know if you're interested in architecture or implementation contribution, I'll be more than happy! 

##Whast is NPOI?

NPOI is a .NET version of POI Java Project. It is an open source .NET library to read and write Microsoft Office file formats formats (*.xls/xlsx, *.doc/docx, *.ppt/pptx). 
You can manually download the repository from [GitHub](https://github.com/nissl-lab/npoi) or install from [NuGet](https://www.nuget.org/packages/NPOI/).

## Installation
* Download the latest version of the LabVIEW NPOI VI Package from the release page.
* Use VIPM to install it on your LabVIEW version (2017+).

## Examples
Below is a simple example to show how to use the LabVIEW NPOI API to interact with Excel or Word document.

![SimpleDocumentCreationExample](https://github.com/Benji667/LabVIEW_NPOI/blob/bcb686f6b338eb219e46d72dd402a0802e551e9f/docs/img/SimpleDocumentCreationExample.png)

## Contributing
Pull requests are welcome. Please use LabVIEW 2017 SP1 to contribute to this repository.

For major changes, please open an issue first to discuss what you would like to change.

# Author
BenjaminR
