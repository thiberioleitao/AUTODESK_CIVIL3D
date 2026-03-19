# Zago Civil 3D Plugin

Exemplo minimo de plugin para Autodesk Civil 3D / AutoCAD 2024+ usando C# e .NET 8 (Windows), compativel com Visual Studio.

## Estrutura do projeto

```text
AUTODESK_CIVIL3D/
|-- src/
|   `-- ZagoCivil3D/
|       |-- HelloCommand.cs
|       `-- ZagoCivil3D.csproj
|-- docs/
|   `-- install.md
`-- README.md
```

## O que este plugin faz

O projeto cria uma Class Library em C# com um comando AutoCAD chamado `HELLO_CIVIL3D`.

Quando o comando e executado dentro do Civil 3D, o plugin escreve a mensagem abaixo na linha de comando:

```text
Hello Civil 3D from plugin
```

## Como compilar no Visual Studio

1. Abra o Visual Studio 2022 ou superior.
2. Abra o arquivo [src/ZagoCivil3D/ZagoCivil3D.csproj](/C:/Users/thib/.codex/worktrees/0bbb/AUTODESK_CIVIL3D/src/ZagoCivil3D/ZagoCivil3D.csproj).
3. Verifique se os caminhos `HintPath` das DLLs da API do AutoCAD apontam para a sua instalacao local do AutoCAD ou Civil 3D 2024+.
4. Compile em `Build > Build Solution`.
5. A DLL sera gerada em `bin/x64/Debug/net8.0-windows/` ou `bin/x64/Release/net8.0-windows/`.

## Como carregar com NETLOAD no Civil 3D

1. Abra o Autodesk Civil 3D.
2. Digite `NETLOAD` na Command Line e pressione Enter.
3. Selecione a DLL compilada `ZagoCivil3D.dll`.
4. Depois do carregamento, digite `HELLO_CIVIL3D`.
5. O Civil 3D mostrara a mensagem `Hello Civil 3D from plugin`.

## Arquivos principais

- [HelloCommand.cs](/C:/Users/thib/.codex/worktrees/0bbb/AUTODESK_CIVIL3D/src/ZagoCivil3D/HelloCommand.cs): implementa o comando com `[CommandMethod]` e inclui comentarios explicando funcoes e variaveis.
- [ZagoCivil3D.csproj](/C:/Users/thib/.codex/worktrees/0bbb/AUTODESK_CIVIL3D/src/ZagoCivil3D/ZagoCivil3D.csproj): define a Class Library e as referencias das DLLs da API do AutoCAD.
- [install.md](/C:/Users/thib/.codex/worktrees/0bbb/AUTODESK_CIVIL3D/docs/install.md): documenta os passos de compilacao e carregamento.

## Observacoes

- O exemplo nao usa dependencias externas.
- O alvo do projeto e `.NET 8 (net8.0-windows)`.
- Se o caminho de instalacao do AutoCAD/Civil 3D for diferente, ajuste o `HintPath` no arquivo do projeto antes de compilar.

