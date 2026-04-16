# AUTODESK_CIVIL3D — Zago Civil 3D Plugin

Plugin C# para Autodesk Civil 3D / AutoCAD 2024+, .NET 8 (Windows).

## Stack
- C# / .NET 8 (net8.0-windows)
- Visual Studio 2022 Community
- API do AutoCAD/Civil 3D (DLLs via HintPath local)
- Dynamo (scripts .dyn externos, Dropbox PADRÃO ZAGO)

## Build
Abrir: src/ZagoCivil3D/ZagoCivil3D.csproj no Visual Studio
Build: Build > Build Solution
Output: bin/x64/Debug/net8.0-windows/ZagoCivil3D.dll

## Como testar
1. Abrir Civil 3D
2. NETLOAD → selecionar a DLL
3. Digitar o comando na Command Line

## Estrutura
- src/ZagoCivil3D/     → código fonte dos comandos
- docs/                → documentação de instalação
- Backup/              → ignorar, artefato do VS
- *.dyn                → scripts Dynamo externos (não editar)

## Convenções do projeto
- Nomes de métodos/variáveis em português
- Prefixo m_ para atributos de classe
- Transações de leitura e escrita separadas

