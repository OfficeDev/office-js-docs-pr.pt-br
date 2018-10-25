---
title: Validar e solucionar problemas com seu manifesto
description: Use estes métodos para validar o manifesto de Suplementos do Office.
ms.date: 12/04/2017
ms.openlocfilehash: 51d644f7cfb7fbad5c9b66be41dc57015202b9be
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639984"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>Validar e solucionar problemas com seu manifesto

Use estes métodos para validar e solucionar problemas no manifesto dos Suplementos do Office. 

- [Validar o manifesto com o Validador de Suplemento do Office](#validate-your-manifest-with-the-office-add-in-validator)   
- [Validar o manifesto com base no esquema XML](#validate-your-manifest-against-the-xml-schema)
- [Usar o log de tempo de execução para depurar o manifesto do suplemento](#use-runtime-logging-to-debug-your-add-in-manifest)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a>Validar o manifesto com o Validador de Suplemento do Office

Para se assegurar de que o arquivo de manifesto que descreve o Suplemento do Office está correto e completo, valide-o com o [Validador de Suplemento do Office](https://github.com/OfficeDev/office-addin-validator).

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a>Para usar o Validador de Suplemento do Office para validar o manifesto

1. Instale o [Node.js](https://nodejs.org/download/). 

2. Abra um prompt de comando/terminal como administrador e instale o Validador de Suplemento do Office e suas dependências globalmente usando o seguinte comando:

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > se já instalou o Yo Office, atualize-o para a última versão e o validador será instalado como uma dependência.

3. Para validar o manifesto, execute o seguinte comando. Substitua MANIFEST.XML pelo caminho para o arquivo XML do manifesto.

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a>Validar o manifesto com base no esquema XML

Para ajudar a garantir que o arquivo de manifesto segue o esquema correto, incluindo quaisquer namespaces para os elementos que você está usando. Se você copiou elementos de outros manifestos de amostra, verifique se também **incluem os namespaces apropriados**. Você pode validar um manifesto com base nos arquivos da [Definição de Esquema XML  (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Pode usar uma ferramenta de validação de esquema XML para executar essa validação. 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>Usar uma ferramenta de validação de esquema XML da linha de comando para validar o manifesto

1.  Instale o [tar](https://www.gnu.org/software/tar/) e o [libxml](http://xmlsoft.org/FAQ.html), caso ainda não tenha instalado.

2.  Execute o comando a seguir. Substitua `XSD_FILE` pelo caminho para o arquivo XSD do manifesto e `XML_FILE` pelo caminho para o arquivo XML do manifesto.
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a>Use o log de tempo de execução para depurar o suplemento 

Você pode usar o log de tempo de execução para depurar o manifesto do suplemento, bem como vários erros de instalação. Esse recurso pode ajudá-lo a identificar e corrigir problemas com o manifesto que não são detectados pela validação do esquema XSD, como uma incompatibilidade entre os IDs do recurso. O log de tempo de execução é particularmente útil para depurar suplementos que implementam comandos de suplementos e funções personalizadas do Excel.   

> [!NOTE]
> O recurso de log de tempo de execução está disponível atualmente para o Office 2016 desktop.

### <a name="to-turn-on-runtime-logging"></a>Para ativar o log de tempo de execução

> [!IMPORTANT]
> O log do tempo de execução afeta o desempenho. Ative-o somente quando precisar depurar problemas com o manifesto do suplemento.

Para ativar o log de tempo de execução:

1. Verifique se você está executando o Office 2016 desktop, build **16.0.7019** ou posterior. 

2. Adicione a chave do registro `RuntimeLogging` em  `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\`. 

3. Defina o valor padrão da chave para o caminho completo do arquivo onde você deseja que o log seja gravado. Para obter um exemplo, veja [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip). 

    > [!NOTE]
    > O diretório no qual o arquivo de log será gravado já deve existir e você deve ter permissões de gravação. 
 
A imagem a seguir mostra como deve ficar o registro. Para desativar o recurso, remova a chave  `RuntimeLogging` do registro. 

![Captura de tela do editor de registro com uma chave de registro RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a>Para solucionar problemas com o manifesto

Para usar o log de tempo de execução para solucionar problemas ao carregar um suplemento:
 
1. [Realize o sideload do suplemento](sideload-office-add-ins-for-testing.md) para testes. 

    > [!NOTE]
    > Recomendamos realizar o sideload apenas do suplemento que você está testando para diminuir a quantidade de mensagens no arquivo de log.

2. Se nada acontece e se você não vê o suplemento (não aparece na caixa de diálogo de suplementos), abra o arquivo de log.

3. Procure no arquivo de log o ID de seu suplemento que foi definido no seu manifesto. No arquivo de log, esse ID está marcado como `SolutionId`. 

No exemplo a seguir, o arquivo de log identifica um controle que aponta para um arquivo de recurso que não existe. Neste exemplo, a correção seria corrigir o erro de digitação no manifesto ou adicionar o recurso ausente.

![Captura de tela de um arquivo de log com uma entrada que especifica a identificação do recurso que não foi encontrado](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>Problemas conhecidos com o log de tempo de execução

Talvez você veja mensagens no arquivo de log que são confusas ou que estão classificadas incorretamente. Por exemplo:

- A mensagem `Medium Current host not in add-in's host list` seguida por `Unexpected Parsed manifest targeting different host` está classificada incorretamente como um erro.

- Se você vir a mensagem `Unexpected Add-in is missing required manifest fields DisplayName` e ela não contiver uma SolutionId, o erro provavelmente não está relacionado ao suplemento que você está depurando. 

- Todas as mensagens `Monitorable` indicam erros esperados do ponto de vista do sistema. Às vezes, indicam um problema com seu manifesto, como um elemento com erro de ortografia que foi ignorado, mas não causou a falha do manifesto. 

## <a name="clear-the-office-cache"></a>Limpar o cache do Office

Se as alterações feitas no manifesto, como nomes de arquivo dos ícones do botão da faixa de opções ou o texto de comandos do suplemento, não entraram em vigor, tente limpar o cache do Office no computador. 

#### <a name="for-windows"></a>No Windows:
Exclua o conteúdo da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>No Mac:
Exclua o conteúdo da pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

#### <a name="for-ios"></a>No iOS:
Chame o `window.location.reload(true)` do JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.

## <a name="see-also"></a>Confira também

- [Manifesto XML dos suplementos do Office](../develop/add-in-manifests.md)
- [Fazer sideload de Suplementos do Office para testes](sideload-office-add-ins-for-testing.md)
- [Depurar suplementos do Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
