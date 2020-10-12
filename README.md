# Scan-and-Clean-Macro-Virus

功能：

扫描office文档中的指定宏病毒，具体扫描规则请自己看代码。C sharp编写，无需依赖，速度贼快。

为什么2020年了还要写代码检测宏？

![](https://github.com/xiaoxiaoleo/Scan-and-Clean-Macro-Virus/raw/main/q.gif)

因为某高级AI人工智能（zhang）杀软不支持扫描宏病毒。





编译：

 mcs /reference:OpenMcdf.dll,System.IO.Compression.FileSystem.dll,System.IO.Compression.dll *.cs

