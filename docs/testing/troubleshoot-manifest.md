---
title: Validar e solucionar problemas com seu manifesto
description: Use estes métodos para validar o manifesto de suplementos do Office
ms.date: 10/29/2019
localization_priority: Priority
ms.openlocfilehash: c1af6308a975bf9204a519e21f828454d286aa19
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924805"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="46b38-103">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="46b38-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="46b38-104">Talvez você queira validar o arquivo de manifesto do seu suplemento para garantir que ele está correto e completo.</span><span class="sxs-lookup"><span data-stu-id="46b38-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="46b38-105">A validação também pode identificar problemas que estejam causando o erro "seu manifesto de suplemento não é válido" quando você tenta realizar o sideload do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="46b38-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="46b38-106">Este artigo descreve várias maneiras de validar o arquivo de manifesto e solucionar problemas com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="46b38-106">This article describes multiple ways to validate the manifest file and troubleshoot problems with your add-in.</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="46b38-107">Validar o manifesto com o gerador Yeoman para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="46b38-107">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="46b38-108">Se você usou o [gerador de Yeoman para suplementos](https://www.npmjs.com/package/generator-office) do Office para criar seu suplemento, você também pode usá-lo para validar o arquivo de manifesto do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="46b38-108">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="46b38-109">Execute o seguinte comando no diretório raiz do seu projeto:</span><span class="sxs-lookup"><span data-stu-id="46b38-109">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![Gif animado que mostra o validador Yo Office em execução na linha de comando e gerando os resultados que mostram que a validação foi aprovada](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="46b38-111">Para ter acesso a essa funcionalidade, o projeto de suplemento deve ter sido criado usando o [Gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office) versão 1.1.17 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="46b38-111">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="46b38-112">Valide seu manifesto com o office-addin-manifest</span><span class="sxs-lookup"><span data-stu-id="46b38-112">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="46b38-113">Se você não tiver usado o [gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office) para criar seu suplemento, você também pode usá-lo para validar o arquivo de manifesto usando o[office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span><span class="sxs-lookup"><span data-stu-id="46b38-113">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="46b38-114">Instale o [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="46b38-114">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="46b38-115">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="46b38-115">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="46b38-116">Substitua o `MANIFEST_FILE` pelo nome do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="46b38-116">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > <span data-ttu-id="46b38-117">Se ao executar esse comando resultar na mensagem de erro "A sintaxe do comando não é válida".</span><span class="sxs-lookup"><span data-stu-id="46b38-117">If running this command results in the error message "The command syntax is not valid."</span></span> <span data-ttu-id="46b38-118">(como o comando `validate` não é reconhecido), execute o seguinte comando para validar o manifesto (substitua o `MANIFEST_FILE` pelo nome do arquivo de manifesto):</span><span class="sxs-lookup"><span data-stu-id="46b38-118">(because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span> 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="46b38-119">Validar seu manifesto em relação ao esquema XML</span><span class="sxs-lookup"><span data-stu-id="46b38-119">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="46b38-120">É possível validar um manifesto em relação aos arquivos de [Definição de esquema XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas).</span><span class="sxs-lookup"><span data-stu-id="46b38-120">You can validate the manifest file against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="46b38-121">Para ajudar a garantir que o arquivo de manifesto segue o esquema correto, incluindo todos os namespaces para os elementos que você está usando.</span><span class="sxs-lookup"><span data-stu-id="46b38-121">This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="46b38-122">Se você copiou elementos de outros manifestos da amostra, verifique se também **incluiu os namespaces apropriados**.</span><span class="sxs-lookup"><span data-stu-id="46b38-122">If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="46b38-123">É possível usar uma ferramenta de validação de esquema XML para executar essa validação.</span><span class="sxs-lookup"><span data-stu-id="46b38-123">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="46b38-124">Para usar uma ferramenta de validação de esquema XML da linha de comando para validar seu manifesto</span><span class="sxs-lookup"><span data-stu-id="46b38-124">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="46b38-125">Instale o [tar](https://www.gnu.org/software/tar/) e o [libxml](http://xmlsoft.org/FAQ.html), caso ainda não tenha o feito.</span><span class="sxs-lookup"><span data-stu-id="46b38-125">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="46b38-p106">Execute o comando a seguir. Substitua `XSD_FILE` pelo caminho para o arquivo XSD do manifesto e `XML_FILE` pelo caminho para o arquivo XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="46b38-p106">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="46b38-128">Usar o log de tempo de execução para depurar seu suplemento</span><span class="sxs-lookup"><span data-stu-id="46b38-128">Use runtime logging to debug your add-in</span></span>

<span data-ttu-id="46b38-129">Você pode usar o log de tempo de execução para depurar o manifesto do seu suplemento, assim como diversos erros de instalação.</span><span class="sxs-lookup"><span data-stu-id="46b38-129">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="46b38-130">Esse recurso pode ajudá-lo a identificar e corrigir problemas com seu manifesto que não são detectados pela validação de esquema XSD, como uma incompatibilidade entre as identificações dos recursos.</span><span class="sxs-lookup"><span data-stu-id="46b38-130">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="46b38-131">O log de tempo de execução é particularmente útil para depurar suplementos que implementam comandos de suplemento e funções personalizadas do Excel.</span><span class="sxs-lookup"><span data-stu-id="46b38-131">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="46b38-132">O recurso de log de tempo de execução está atualmente disponível para o Office 2016 para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="46b38-132">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="46b38-133">o log do tempo de execução afeta o desempenho.</span><span class="sxs-lookup"><span data-stu-id="46b38-133">Runtime Logging affects performance.</span></span> <span data-ttu-id="46b38-134">Ative-o somente quando precisar depurar problemas com o manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="46b38-134">Turn it on only when you need to debug issues with your add-in manifest.</span></span>

### <a name="use-runtime-logging-from-the-command-line"></a><span data-ttu-id="46b38-135">Use o log de tempo de execução na linha de comandos</span><span class="sxs-lookup"><span data-stu-id="46b38-135">Use runtime logging from the command line</span></span>

<span data-ttu-id="46b38-136">Habilitar o log de tempo de execução na linha de comando é a maneira mais rápida de usar essa ferramenta de log.</span><span class="sxs-lookup"><span data-stu-id="46b38-136">Enabling runtime logging from the command line is the fastest way to use this logging tool.</span></span> <span data-ttu-id="46b38-137">Esse usa npx, que é fornecido por padrão como parte do npm@5.2.0+.</span><span class="sxs-lookup"><span data-stu-id="46b38-137">These use npx, which is provided by default as part of npm@5.2.0+.</span></span> <span data-ttu-id="46b38-138">Se você possui uma versão anterior do [npm](https://www.npmjs.com/), tente as instruções do [Log de tempo de execução no Windows](#runtime-logging-on-windows) ou do [Log de tempo de execução no Mac](#runtime-logging-on-mac), ou [instale o npx](https://www.npmjs.com/package/npx).</span><span class="sxs-lookup"><span data-stu-id="46b38-138">If you have an earlier version of [npm](https://www.npmjs.com/), try [Runtime logging on Windows](#runtime-logging-on-windows) or [Runtime logging on Mac](#runtime-logging-on-mac) instructions, or [install npx](https://www.npmjs.com/package/npx).</span></span>

- <span data-ttu-id="46b38-139">Para habilitar o log de tempo de execução:</span><span class="sxs-lookup"><span data-stu-id="46b38-139">To enable AD FS logging</span></span>
    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```
- <span data-ttu-id="46b38-140">Para habilitar o log de tempo de execução apenas para um arquivo específico, use o mesmo comando com um nome de arquivo:</span><span class="sxs-lookup"><span data-stu-id="46b38-140">To enable runtime logging only for a specific file, use the same command with a filename:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- <span data-ttu-id="46b38-141">Para desabilitar o log de tempo de execução:</span><span class="sxs-lookup"><span data-stu-id="46b38-141">To disable runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- <span data-ttu-id="46b38-142">Para exibir se o log de tempo de execução está ativado:</span><span class="sxs-lookup"><span data-stu-id="46b38-142">To display whether runtime logging is enabled:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- <span data-ttu-id="46b38-143">Para exibir ajuda na linha de comandos para o log de tempo de execução:</span><span class="sxs-lookup"><span data-stu-id="46b38-143">To display help within the command line for runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

### <a name="runtime-logging-on-windows"></a><span data-ttu-id="46b38-144">Log de tempo de execução no Windows</span><span class="sxs-lookup"><span data-stu-id="46b38-144">Runtime logging on Windows</span></span>

1. <span data-ttu-id="46b38-145">Verifique se você está executando o Office 2016 para área de trabalho na compilação **16.0.7019** ou posterior.</span><span class="sxs-lookup"><span data-stu-id="46b38-145">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="46b38-146">Adicione a chave do registro `RuntimeLogging` em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`</span><span class="sxs-lookup"><span data-stu-id="46b38-146">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="46b38-147">Se a chave (pasta) `Developer` ainda não existir em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, conclua as seguintes etapas para criá-la:</span><span class="sxs-lookup"><span data-stu-id="46b38-147">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="46b38-148">Clique com o botão direito do mouse na chave (pasta) **WEF** e selecione **Novo** > **Chave**.</span><span class="sxs-lookup"><span data-stu-id="46b38-148">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="46b38-149">Nomeie a nova chave como **Developer**.</span><span class="sxs-lookup"><span data-stu-id="46b38-149">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="46b38-150">Defina o valor padrão da chave **RuntimeLogging** para o caminho completo do arquivo em que você deseja que o log seja gravado.</span><span class="sxs-lookup"><span data-stu-id="46b38-150">Set the default value of the key to the full path of the file where you want the log to be written.</span></span> <span data-ttu-id="46b38-151">Para obter um exemplo, veja [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="46b38-151">For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="46b38-152">A pasta na qual o arquivo de log será gravado deverá existir e você precisará ter permissões de gravação.</span><span class="sxs-lookup"><span data-stu-id="46b38-152">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="46b38-p111">A imagem a seguir mostra qual deve ser a aparência do registro. Para desativar o recurso, remova a chave do registro `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="46b38-p111">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Captura de tela do editor do registro com uma chave de registro RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

### <a name="runtime-logging-on-mac"></a><span data-ttu-id="46b38-156">Log de tempo de execução no Mac</span><span class="sxs-lookup"><span data-stu-id="46b38-156">Runtime logging on Mac</span></span>

1. <span data-ttu-id="46b38-157">Verifique se você está executando o build de área de trabalho do Office 2016 **16.27** (19071500) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="46b38-157">Make sure that you are running Office 2016 desktop build **16.27** (19071500) or later.</span></span>

2. <span data-ttu-id="46b38-158">Abra o **Terminal** e defina uma preferência de log de tempo de execução usando o comando `defaults`:</span><span class="sxs-lookup"><span data-stu-id="46b38-158">Open **Terminal** and set a runtime logging preference by using the `defaults` command:</span></span>
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    <span data-ttu-id="46b38-159">`<bundle id>` identifica quais hosts devem ser habilitados no log de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="46b38-159">`<bundle id>` identifies which the host for which to enable runtime logging.</span></span> <span data-ttu-id="46b38-160">`<file_name>` é o nome do arquivo de texto no qual o log será gravado.</span><span class="sxs-lookup"><span data-stu-id="46b38-160">`<file_name>` is the name of the text file to which the log will be written.</span></span>

    <span data-ttu-id="46b38-161">Defina `<bundle id>` para um dos seguintes valores para habilitar o log de tempo de execução do host correspondente:</span><span class="sxs-lookup"><span data-stu-id="46b38-161">Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding host:</span></span>

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

<span data-ttu-id="46b38-162">O exemplo a seguir habilita o log de tempo de execução do Word e, em seguida, abre o arquivo de log:</span><span class="sxs-lookup"><span data-stu-id="46b38-162">The following example enables runtime logging for Word and then opens the log file:</span></span>

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> <span data-ttu-id="46b38-163">Será preciso reiniciar o Office depois de executar o comando `defaults` para habilitar o log de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="46b38-163">You'll need to restart Office after running the `defaults` command to enable runtime logging.</span></span>

<span data-ttu-id="46b38-164">Para desativar o log de tempo de execução, use o comando `defaults delete`:</span><span class="sxs-lookup"><span data-stu-id="46b38-164">To turn off runtime logging, use the `defaults delete` command:</span></span>

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

<span data-ttu-id="46b38-165">O exemplo a seguir desabilitará o log de tempo de execução do Word:</span><span class="sxs-lookup"><span data-stu-id="46b38-165">The following example will turn off runtime logging for Word:</span></span>

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="46b38-166">Para solucionar problemas com o manifesto</span><span class="sxs-lookup"><span data-stu-id="46b38-166">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="46b38-167">Para usar o log do tempo de execução para solucionar problemas ao carregar um suplemento:</span><span class="sxs-lookup"><span data-stu-id="46b38-167">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="46b38-168">[Realize o sideload do seu suplemento](sideload-office-add-ins-for-testing.md) para teste.</span><span class="sxs-lookup"><span data-stu-id="46b38-168">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="46b38-169">Recomendamos realizar o sideload apenas do suplemento que você está testando para minimizar a quantidade de mensagens no arquivo de log.</span><span class="sxs-lookup"><span data-stu-id="46b38-169">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="46b38-170">Se nada acontecer e você não vir seu suplemento (e ele não estiver aparecendo na caixa de diálogo de suplementos), abra o arquivo de log.</span><span class="sxs-lookup"><span data-stu-id="46b38-170">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="46b38-p113">Procure pela ID de seu suplemento no arquivo de log, definida no seu manifesto. No arquivo de log, essa ID está marcada como `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="46b38-p113">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="46b38-p114">No exemplo a seguir, o arquivo de log identifica um controle que aponta para um arquivo de recurso que não existe. Neste exemplo, a correção seria reparar o erro de digitação no manifesto ou adicionar o recurso que está faltando.</span><span class="sxs-lookup"><span data-stu-id="46b38-p114">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Captura de tela de um arquivo de log com uma entrada que especifica uma identificação de recurso que não foi encontrado](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="46b38-176">Problemas conhecidos com o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="46b38-176">Known issues with runtime logging</span></span>

<span data-ttu-id="46b38-p115">Talvez você veja mensagens no arquivo de log que são confusas ou que estão classificadas incorretamente. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="46b38-p115">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="46b38-179">A mensagem `Medium Current host not in add-in's host list` seguida por `Unexpected Parsed manifest targeting different host` é incorretamente classificada como um erro.</span><span class="sxs-lookup"><span data-stu-id="46b38-179">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="46b38-180">Se você vir a mensagem `Unexpected Add-in is missing required manifest fields DisplayName` e ela não contiver uma SolutionId, o erro provavelmente não está relacionado ao suplemento que você está depurando.</span><span class="sxs-lookup"><span data-stu-id="46b38-180">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="46b38-p116">Todas as mensagens `Monitorable` indicam erros esperados do ponto de vista do sistema. Às vezes, indicam um problema com o seu manifesto, como um elemento que foi soletrado incorretamente e que foi ignorado, mas que não fez com que o manifesto falhasse.</span><span class="sxs-lookup"><span data-stu-id="46b38-p116">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="46b38-183">Limpar o cache do Office</span><span class="sxs-lookup"><span data-stu-id="46b38-183">Clear the Office cache</span></span>

<span data-ttu-id="46b38-184">Se alterações feitas no manifesto, como nomes de arquivo de ícones de botão da faixa de opções ou texto de comandos de suplemento, não parecerem entrar em vigor, experimente limpar o cache do Office no computador.</span><span class="sxs-lookup"><span data-stu-id="46b38-184">If changes you've made in the manifest, such as file names of ribbon button icons or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="46b38-185">No Windows:</span><span class="sxs-lookup"><span data-stu-id="46b38-185">For Windows:</span></span>
<span data-ttu-id="46b38-186">Exclua os conteúdos da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="46b38-186">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="46b38-187">Para Mac:</span><span class="sxs-lookup"><span data-stu-id="46b38-187">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="46b38-188">No iOS:</span><span class="sxs-lookup"><span data-stu-id="46b38-188">For iOS:</span></span>
<span data-ttu-id="46b38-p117">Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="46b38-p117">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="46b38-191">Confira também</span><span class="sxs-lookup"><span data-stu-id="46b38-191">See also</span></span>

- [<span data-ttu-id="46b38-192">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="46b38-192">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="46b38-193">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="46b38-193">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="46b38-194">Depurar suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="46b38-194">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
