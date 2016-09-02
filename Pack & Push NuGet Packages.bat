cd /d %cd%
rmdir /s /q NuGet
mkdir NuGet

.nuget\nuget update -Self
.nuget\nuget pack CsvHelperAsync\CsvHelperAsync.csproj -Prop Configuration=Release -OutputDirectory NuGet -Build
.nuget\nuget push NuGet\*.nupkg -Source https://www.nuget.org/api/v2/package

pause
