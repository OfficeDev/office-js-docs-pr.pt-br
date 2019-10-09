---
title: Validar e solucionar problemas com seu manifesto
description: Use estes métodos para validar o manifesto de suplementos do Office
ms.date: 09/18/2019
localization_priority: Priority
ms.openlocfilehash: c320c05b944bba9e24a4d3c0e5ef514ac13cc3c6
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035333"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>Validar e solucionar problemas com seu manifesto

Talvez você queira validar o arquivo de manifesto do seu suplemento para garantir que ele está correto e completo. A validação também pode identificar problemas que estejam causando o erro "seu manifesto de suplemento não é válido" quando você tenta realizar o sideload do seu suplemento. Este artigo descreve várias maneiras de validar o arquivo de manifesto e solucionar problemas com o suplemento.

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>Validar o manifesto com o gerador Yeoman para Suplementos do Office

Se você usou o [gerador de Yeoman para suplementos](https://www.npmjs.com/package/generator-office) do Office para criar seu suplemento, você também pode usá-lo para validar o arquivo de manifesto do seu projeto. Execute o seguinte comando no diretório raiz do seu projeto:

```command&nbsp;line
npm run validate
```

![Gif animado que mostra o validador Yo Office em execução na linha de comando e gerando os resultados que mostram que a validação foi aprovada](../images/yo-office-validator.gif)

> [!NOTE]
> Para ter acesso a essa funcionalidade, o projeto de suplemento deve ter sido criado usando o [Gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office) versão 1.1.17 ou posterior.

## <a name="validate-your-manifest-with-office-addin-manifest"></a>Valide seu manifesto com o office-addin-manifest

Se você não tiver usado o [gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office) para criar seu suplemento, você também pode usá-lo para validar o arquivo de manifesto usando o[office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).

1. Instale o [Node.js](https://nodejs.org/download/).

2. Execute o seguinte comando no diretório raiz do seu projeto. Substitua o `MANIFEST_FILE` pelo nome do arquivo de manifesto.

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > Se ao executar esse comando resultar na mensagem de erro "A sintaxe do comando não é válida". (como o comando `validate` não é reconhecido), execute o seguinte comando para validar o manifesto (substitua o `MANIFEST_FILE` pelo nome do arquivo de manifesto): 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a>Validar seu manifesto em relação ao esquema XML

É possível validar um manifesto em relação aos arquivos de [Definição de esquema XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Para ajudar a garantir que o arquivo de manifesto segue o esquema correto, incluindo todos os namespaces para os elementos que você está usando. Se você copiou elementos de outros manifestos da amostra, verifique se também **incluiu os namespaces apropriados**. É possível usar uma ferramenta de validação de esquema XML para executar essa validação.

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>Para usar uma ferramenta de validação de esquema XML da linha de comando para validar seu manifesto

1. Instale o [tar](https://www.gnu.org/software/tar/) e o [libxml](http://xmlsoft.org/FAQ.html), caso ainda não tenha o feito.

2. Execute o comando a seguir. Substitua `XSD_FILE` pelo caminho para o arquivo XSD do manifesto e `XML_FILE` pelo caminho para o arquivo XML do manifesto.
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a>Usar o log de tempo de execução para depurar seu suplemento

Você pode usar o log de tempo de execução para depurar o manifesto do seu suplemento, assim como diversos erros de instalação. Esse recurso pode ajudá-lo a identificar e corrigir problemas com seu manifesto que não são detectados pela validação de esquema XSD, como uma incompatibilidade entre as identificações dos recursos. O log de tempo de execução é particularmente útil para depurar suplementos que implementam comandos de suplemento e funções personalizadas do Excel.   

> [!NOTE]
> O recurso de log de tempo de execução está atualmente disponível para o Office 2016 para área de trabalho.

> [!IMPORTANT]
> o log do tempo de execução afeta o desempenho. Ative-o somente quando precisar depurar problemas com o manifesto do suplemento.

### <a name="runtime-logging-on-windows"></a>Log de tempo de execução no Windows

1. Verifique se você está executando o Office 2016 para área de trabalho na compilação **16.0.7019** ou posterior. 

2. Adicione a chave do registro `RuntimeLogging` em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\` 

    > [!NOTE]
    > Se a chave (pasta) `Developer` ainda não existir em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, conclua as seguintes etapas para criá-la: 
    > 1. Clique com o botão direito do mouse na chave (pasta) **WEF** e selecione **Novo** > **Chave**.
    > 2. Nomeie a nova chave como **Developer**.

3. Defina o valor padrão da chave para o caminho completo do arquivo onde você deseja que o log seja gravado. Para obter um exemplo, veja [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip). 

    > [!NOTE]
    > A pasta na qual o arquivo de log será gravado deverá existir e você precisará ter permissões de gravação. 
 
A imagem a seguir mostra qual deve ser a aparência do registro. Para desativar o recurso, remova a chave do registro `RuntimeLogging`. 

![Captura de tela do editor do registro com uma chave de registro RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

### <a name="runtime-logging-on-mac"></a>Log de tempo de execução no Mac

1. Verifique se você está executando o build de área de trabalho do Office 2016 **16.27** (19071500) ou posterior.

2. Abra o **Terminal** e defina uma preferência de log de tempo de execução usando o comando `defaults`:
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    `<bundle id>` identifica quais hosts devem ser habilitados no log de tempo de execução. `<file_name>` é o nome do arquivo de texto no qual o log será gravado.

    Defina `<bundle id>` para um dos seguintes valores para habilitar o log de tempo de execução do host correspondente:

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

O exemplo a seguir habilita o log de tempo de execução do Word e, em seguida, abre o arquivo de log:

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> Será preciso reiniciar o Office depois de executar o comando `defaults` para habilitar o log de tempo de execução.

Para desativar o log de tempo de execução, use o comando `defaults delete`:

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

O exemplo a seguir desabilitará o log de tempo de execução do Word:

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

### <a name="to-troubleshoot-issues-with-your-manifest"></a>Para solucionar problemas com o manifesto

Para usar o log do tempo de execução para solucionar problemas ao carregar um suplemento:
 
1. [Realize o sideload do seu suplemento](sideload-office-add-ins-for-testing.md) para teste. 

    > [!NOTE]
    > Recomendamos realizar o sideload apenas do suplemento que você está testando para minimizar a quantidade de mensagens no arquivo de log.

2. Se nada acontecer e você não vir seu suplemento (e ele não estiver aparecendo na caixa de diálogo de suplementos), abra o arquivo de log.

3. Procure pela ID de seu suplemento no arquivo de log, definida no seu manifesto. No arquivo de log, essa ID está marcada como `SolutionId`. 

No exemplo a seguir, o arquivo de log identifica um controle que aponta para um arquivo de recurso que não existe. Neste exemplo, a correção seria reparar o erro de digitação no manifesto ou adicionar o recurso que está faltando.

![Captura de tela de um arquivo de log com uma entrada que especifica uma identificação de recurso que não foi encontrado](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>Problemas conhecidos com o log de tempo de execução

Talvez você veja mensagens no arquivo de log que são confusas ou que estão classificadas incorretamente. Por exemplo:

- A mensagem `Medium Current host not in add-in's host list` seguida por `Unexpected Parsed manifest targeting different host` é incorretamente classificada como um erro.

- Se você vir a mensagem `Unexpected Add-in is missing required manifest fields DisplayName` e ela não contiver uma SolutionId, o erro provavelmente não está relacionado ao suplemento que você está depurando. 

- Todas as mensagens `Monitorable` indicam erros esperados do ponto de vista do sistema. Às vezes, indicam um problema com o seu manifesto, como um elemento que foi soletrado incorretamente e que foi ignorado, mas que não fez com que o manifesto falhasse. 

## <a name="clear-the-office-cache"></a>Limpar o cache do Office

Se alterações feitas no manifesto, como nomes de arquivo de ícones de botão da faixa de opções ou texto de comandos de suplemento, não parecerem entrar em vigor, experimente limpar o cache do Office no computador. 

#### <a name="for-windows"></a>No Windows:
Exclua os conteúdos da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>Para Mac:

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>No iOS:
Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.

## <a name="see-also"></a>Confira também

- [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md)
- [Realizar sideload de suplementos do Office para teste](sideload-office-add-ins-for-testing.md)
- [Depurar suplementos do Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
