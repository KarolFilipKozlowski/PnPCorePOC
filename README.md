# PnPCorePOC by Karol Kozłowski
It's my repository helper for [Microsoft 365 & Power Platform Community](https://pnp.github.io/).

### Resources:

🪪 See my blog: [Karol Kozłowski I CitDev](https://citdev.pl/blog/)

🔗PnP Core Documentation: [# PnP Core SDK](https://pnp.github.io/pnpcore/)

# Start Here:

## 1) Create new Solution:

Add new project (.NET Framework) with .NET Framework 4.8
From NuGet add:
- **Microsoft.Extensions.Hosting** package
- **PnP.Core.Auth** package
- **GitInfo** package

**Add project to a code repository!**

## 2) Edit _AssemblyInfo.cs_, add on end:
Edit `AssemblyCopyrightt`:
 ````
[assembly: AssemblyProduct("https://github.com/KarolFilipKozlowski/PnPCorePOC | " + ThisAssembly.Git.Branch)]
[assembly: AssemblyCopyright("CitDeV ©  2023")]
````
On end change/add:
````
[assembly: AssemblyVersion(ThisAssembly.Git.BaseVersion.Major + "." + ThisAssembly.Git.BaseVersion.Minor + "." + ThisAssembly.Git.BaseVersion.Patch)]

[assembly: AssemblyFileVersion(ThisAssembly.Git.SemVer.Major + "." + ThisAssembly.Git.SemVer.Minor + "." + ThisAssembly.Git.SemVer.Patch)]

[assembly: AssemblyInformationalVersion(
    ThisAssembly.Git.SemVer.Major + "." +
    ThisAssembly.Git.SemVer.Minor + "." +
    ThisAssembly.Git.Commits + "-" +
    ThisAssembly.Git.Commit)]
````

## 3) Go to _properties_ of app:
- Select **Application**. In Icon and manifest add icon: **PnPCoreApp_254.ico**.
- Select **Build**. In platform target select: **x64**.
- Select **Build Events**. In pre-build event add.
````
rmdir $(SolutionDir)APP /s /q

mkdir $(SolutionDir)APP
mkdir $(SolutionDir)APP\Bin
mkdir $(SolutionDir)APP\Logs

copy $(TargetDir)*.exe  $(SolutionDir)APP
copy $(TargetDir)*.config  $(SolutionDir)APP
copy $(TargetDir)*.dll  $(SolutionDir)APP\Bin

powershell Compress-Archive -Path '$(SolutionDir)APP' -DestinationPath '$(SolutionDir)APP\$(ProjectName).zip' -Force
````

## 4a) Add _AppSecret.config_:
Select project Add new item -> Application Configuration File, add **AppSecret.config**.
Replace code with:
````
﻿<appSettings>
	<add key="SPO:0:TenantID" value="8ae35f9e-b3c6-486b-bdf8-5c8da0cff7b9"/>
	<add key="SPO:0:SiteURL" value="https://contoso.sharepoint.com/sites/pnpdemo"/>
	<add key="SPO:0:ApplicationID" value="64f4c29a-b1d5-47c5-91d3-ecdfaedeb72a"/>
	<add key="SPO:0:CertificateThumbprint " value="4AA402CDC596471696C6159254DEF6B30ABBB44D"/>
</appSettings>
````
**In git add this file for ignore!**

## 4b) Edit _App.config_:
In end of **<configuration>** add:
````
<appSettings file="AppSecret.config">
</appSettings>
````
After **<assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">** add:
````
<probing privatePath="Bin"/>
````
