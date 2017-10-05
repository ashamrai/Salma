ECHO Preparing solution for installer creation...
del nsis\SALMA.msi
del nsis\Setup.exe 
del Build\Setup.exe
RMDIR Build
mkdir Build
ECHO Building WiX project... 
ECHO Creating Product.wixobj...
ThirdPartyTools\Wix\candle.exe -drelease -dcodepage=1252 -d"DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\\" -dSolutionDir= -dSolutionExt=.sln -dSolutionFileName=Salma2010.sln -dSolutionName=Salma2010 -dSolutionPath=Salma2010.sln -dConfiguration=release -dOutDir=bin\release\ -dPlatform=x86 -dProjectDir=SalmaInstaller\ -dProjectExt=.wixproj -dProjectFileName=SalmaInstaller.wixproj -dProjectName=SalmaInstaller -dProjectPath=SalmaInstaller\SalmaInstaller.wixproj -dTargetDir=SalmaInstaller\bin\release\ -dTargetExt=.msi -dTargetFileName=SalmaInstaller.msi -dTargetName=SalmaInstaller -dTargetPath=SalmaInstaller\bin\release\SalmaInstaller.msi -dSalma.Configuration=release -d"Salma.FullConfiguration=release|AnyCPU" -dSalma.Platform=AnyCPU -dSalma.ProjectDir=Salma2010\ -dSalma.ProjectExt=.csproj -dSalma.ProjectFileName=Salma.csproj -dSalma.ProjectName=Salma -dSalma.ProjectPath=Salma2010\Salma.csproj -dSalma.TargetDir=Salma2010\bin\release\ -dSalma.TargetExt=.dll -dSalma.TargetFileName=Salma.dll -dSalma.TargetName=Salma -dSalma.TargetPath=Salma2010\bin\release\Salma.dll -out SalmaInstaller\obj\release\ -arch x86 -ext "ThirdPartyTools\Wix\WixUIExtension.dll" -ext "ThirdPartyTools\Wix\WixNetFxExtension.dll" SalmaInstaller\Product.wxs
ECHO Product.wixobj created
ECHO Creating SalmaInstaller
ThirdPartyTools\Wix\Light.exe -out nsis\SALMA.msi -pdbout nsis\Salma.wixpdb -cultures:en-us -ext "ThirdPartyTools\Wix\WixUIExtension.dll" -ext "ThirdPartyTools\Wix\WixNetFxExtension.dll" -contentsfile SalmaInstaller\obj\release\SalmaInstaller.wixproj.BindContentsFileListru-ru.txt -outputsfile SalmaInstaller\obj\release\SalmaInstaller.wixproj.BindOutputsFileListru-ru.txt -builtoutputsfile SalmaInstaller\obj\release\SalmaInstaller.wixproj.BindBuiltOutputsFileListru-ru.txt -wixprojectfile SalmaInstaller\SalmaInstaller.wixproj SalmaInstaller\obj\release\Product.wixobj"
ECHO Created SalmaInstaller
del nsis\Salma.wixpdb
::NSIS compile command
ThirdPartyTools\NSIS\makensis.exe nsis\SALMA.nsi
copy "nsis\Setup.exe" "Build\Setup.exe"
del nsis\SALMA.msi


del nsis\Setup.exe 