# Instalacao do plugin

Este exemplo foi pensado para Civil 3D / AutoCAD 2024 ou superior.

## 1. Ajustar referencias da API

Antes de compilar, confirme que os caminhos das DLLs no arquivo `src/ZagoCivil3D/ZagoCivil3D.csproj` apontam para a instalacao local do AutoCAD ou Civil 3D:

- `AcCoreMgd.dll`
- `AcDbMgd.dll`
- `AcMgd.dll`

Se o produto estiver instalado em outra pasta, atualize os valores de `HintPath`.

## 2. Compilar no Visual Studio

1. Abra a pasta do projeto no Visual Studio 2022 ou superior.
2. Abra o arquivo `src/ZagoCivil3D/ZagoCivil3D.csproj`.
3. Restaure e compile o projeto em `Build > Build Solution`.
4. O arquivo DLL sera gerado em `src/ZagoCivil3D/bin/x64/Debug/net8.0-windows/ZagoCivil3D.dll` ou `src/ZagoCivil3D/bin/x64/Release/net8.0-windows/ZagoCivil3D.dll`.

## 3. Carregar no Civil 3D

1. Abra o Autodesk Civil 3D.
2. Digite `NETLOAD` na linha de comando.
3. Selecione a DLL compilada `ZagoCivil3D.dll`.
4. Depois do carregamento, digite `HELLO_CIVIL3D`.
5. O Civil 3D exibira a mensagem `Hello Civil 3D from plugin` na Command Line.

