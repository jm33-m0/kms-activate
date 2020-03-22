# kms-activate
Microsoft Windows/Office 一键激活工具，基于kms.jm33.me的KMS服务器

## NOTE:

- To use this tool, `kms.jm33.me` or your custom KMS server must be reachable
- Change KMS server if activation keeps failing
- Check `Show debug info` to see what happens during the activation process
- Make sure kms server is reachable, and is present in your slmgr.vbs
- [Office 2019 Vol](https://github.com/jm33-m0/kms-activate#office-deployment-tool) is recommended
- Download from [here](https://github.com/jm33-m0/kms-activate/releases)

![screenshot](/img/win-activate.JPG)

## supported products

- [Windows 10 Professional/Enterprise](https://www.microsoft.com/en-us/software-download/windows10)
- [Windows Server 2008 R2, 2012 R2, 2016, 2019 Standard/Datacenter](https://www.microsoft.com/en-us/evalcenter/evaluate-windows-server-2019?filetype=ISO)
- ~~Windows 7~~ (EOL), 8.1 Professional/Enterprise
- [Office 2010, 2013, 2016, 2019](https://github.com/jm33-m0/kms-activate#office-deployment-tool)
- Visio Pro, Project Pro (included in Office Enterprise)

## Office Deployment Tool

you can use [Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117) to download and install Office (VOL) from Microsoft

### modify configuration XML

upon successful extraction, you will see some XML files under office deployment directory,
the one we care about is `configuration-Office2019Enterprise.xml`:

```xml
<Configuration>

  <Add OfficeClientEdition="64" Channel="PerpetualVL2019">
    <Product ID="ProPlus2019Volume">
      <Language ID="en-us" />
    </Product>
    <Product ID="VisioPro2019Volume">
      <Language ID="en-us" />
    </Product>
    <Product ID="ProjectPro2019Volume">
      <Language ID="en-us" />
    </Product>
  </Add>

  <!--  <RemoveMSI All="True" /> -->

  <!--  <Display Level="None" AcceptEULA="TRUE" />  -->

  <!--  <Property Name="AUTOACTIVATE" Value="1" />  -->
</Configuration>
```

comment out the products you don't need, then save a copy as `office.xml`

### how to install

open a powershell window in current directory

```powershell
# download from Microsoft
.\setup.exe /download office.xml
# install when you have finished downloading
.\setup.exe /configure office.xml
```
