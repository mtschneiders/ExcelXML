
# SimpleXL 

Simple .NET library to export Excel (xlsx) files focused on small memory footprint and performance.

[![NuGet](https://img.shields.io/nuget/v/SimpleXL.svg)](https://www.nuget.org/packages/SimpleXL/)
[![Build status](https://img.shields.io/appveyor/ci/mtschneiders/simplexl/master.svg?label=appveyor)](https://ci.appveyor.com/project/mtschneiders/simplexl/branch/master)


## Example
```
using (var file = new XLFile())
{
    file.ConfigureRange("A1:A2", new XLRangeConfig { Font = XLRangeFont.Bold, Border = true });
    file.ConfigureRange("C1:C2", new XLRangeConfig { Format = XLRangeFormat.Percent });
    
    IEnumerable<List<object>> data = GetData();
    file.WriteData(data);
    file.SaveAs(basePath + ".xlsx");
}
```

## Benchmark
``` ini

BenchmarkDotNet=v0.10.8, OS=Windows 10 Redstone 1 (10.0.14393)
Processor=Intel Core i5-4690 CPU 3.50GHz (Haswell), ProcessorCount=4
Frequency=3410075 Hz, Resolution=293.2487 ns, Timer=TSC
  [Host]   : Clr 4.0.30319.42000, 32bit LegacyJIT-v4.6.1648.0
  ShortRun : Clr 4.0.30319.42000, 32bit LegacyJIT-v4.6.1648.0

Job=ShortRun  LaunchCount=1  TargetCount=3  
WarmupCount=3  

```
 |   Method | NumRecords | StringCols | NumberCols |       Mean |      Allocated |
 |--------- |----------- |----------------- |----------------- |-----------:|----------:|
 | **SimpleXL** |      **10000** |               **10** |               **10** |   **411.5 ms** |  **11.23 MB** |
 |   EPPlus |      10000 |               10 |               10 |   622.1 ms |   77.21 MB |
 | **SimpleXL** |     **100000** |               **10** |               **10** | **3,079.6 ms** |  **95.35 MB** |
 |   EPPlus |     100000 |               10 |               10 | 6,229.4 ms | 531.65 MB |
