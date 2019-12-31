---
title: Validar o manifesto de suplemento do Office
description: Saiba como validar um manifesto de suplemento do Office usando o esquema XML e outras ferramentas.
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 09b5841a0180d8cb730ec8b479df1386a0749b60
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/31/2019
ms.locfileid: "40914899"
---
# <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="8d405-103">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="8d405-103">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="8d405-104">Talvez você queira validar o arquivo de manifesto do seu suplemento para garantir que ele está correto e completo.</span><span class="sxs-lookup"><span data-stu-id="8d405-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="8d405-105">A validação também pode identificar problemas que estejam causando o erro "seu manifesto de suplemento não é válido" quando você tenta realizar o sideload do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="8d405-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="8d405-106">Este artigo descreve várias maneiras de validar o arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="8d405-106">This article describes multiple ways to validate the manifest file and troubleshoot problems with your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="8d405-107">Para saber mais sobre como usar o log de tempo de execução para solucionar problemas no manifesto de suplemento, confira [Depurar seu suplemento com o log de tempo de execução](runtime-logging.md).</span><span class="sxs-lookup"><span data-stu-id="8d405-107">For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="8d405-108">Validar o manifesto com o gerador Yeoman para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8d405-108">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="8d405-109">Se você usou o [gerador de Yeoman para suplementos](https://www.npmjs.com/package/generator-office) do Office para criar seu suplemento, você também pode usá-lo para validar o arquivo de manifesto do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="8d405-109">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="8d405-110">Execute o seguinte comando no diretório raiz do seu projeto:</span><span class="sxs-lookup"><span data-stu-id="8d405-110">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![Gif animado que mostra o validador Yo Office em execução na linha de comando e gerando os resultados que mostram que a validação foi aprovada](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="8d405-112">Para ter acesso a essa funcionalidade, o projeto de suplemento deve ter sido criado usando o [Gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office) versão 1.1.17 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="8d405-112">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="8d405-113">Valide seu manifesto com o office-addin-manifest</span><span class="sxs-lookup"><span data-stu-id="8d405-113">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="8d405-114">Se você não tiver usado o [gerador Yeoman para Suplementos do Office](https://www.npmjs.com/package/generator-office) para criar seu suplemento, você também pode usá-lo para validar o arquivo de manifesto usando o[office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span><span class="sxs-lookup"><span data-stu-id="8d405-114">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="8d405-115">Instale o [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="8d405-115">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="8d405-116">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="8d405-116">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="8d405-117">Substitua o `MANIFEST_FILE` pelo nome do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="8d405-117">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > <span data-ttu-id="8d405-118">Se ao executar esse comando resultar na mensagem de erro "A sintaxe do comando não é válida".</span><span class="sxs-lookup"><span data-stu-id="8d405-118">If running this command results in the error message "The command syntax is not valid."</span></span> <span data-ttu-id="8d405-119">(como o comando `validate` não é reconhecido), execute o seguinte comando para validar o manifesto (substitua o `MANIFEST_FILE` pelo nome do arquivo de manifesto):</span><span class="sxs-lookup"><span data-stu-id="8d405-119">(because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span> 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="8d405-120">Validar seu manifesto em relação ao esquema XML</span><span class="sxs-lookup"><span data-stu-id="8d405-120">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="8d405-121">É possível validar um manifesto em relação aos arquivos de [Definição de esquema XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas).</span><span class="sxs-lookup"><span data-stu-id="8d405-121">You can validate the manifest file against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="8d405-122">Para ajudar a garantir que o arquivo de manifesto segue o esquema correto, incluindo todos os namespaces para os elementos que você está usando.</span><span class="sxs-lookup"><span data-stu-id="8d405-122">This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="8d405-123">Se você copiou elementos de outros manifestos da amostra, verifique se também **incluiu os namespaces apropriados**.</span><span class="sxs-lookup"><span data-stu-id="8d405-123">If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="8d405-124">É possível usar uma ferramenta de validação de esquema XML para executar essa validação.</span><span class="sxs-lookup"><span data-stu-id="8d405-124">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="8d405-125">Para usar uma ferramenta de validação de esquema XML da linha de comando para validar seu manifesto</span><span class="sxs-lookup"><span data-stu-id="8d405-125">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="8d405-126">Instale o [tar](https://www.gnu.org/software/tar/) e o [libxml](http://xmlsoft.org/FAQ.html), caso ainda não tenha o feito.</span><span class="sxs-lookup"><span data-stu-id="8d405-126">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="8d405-p106">Execute o comando a seguir. Substitua `XSD_FILE` pelo caminho para o arquivo XSD do manifesto e `XML_FILE` pelo caminho para o arquivo XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="8d405-p106">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a><span data-ttu-id="8d405-129">Confira também</span><span class="sxs-lookup"><span data-stu-id="8d405-129">See also</span></span>

- [<span data-ttu-id="8d405-130">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8d405-130">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="8d405-131">Limpar o cache do Office</span><span class="sxs-lookup"><span data-stu-id="8d405-131">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="8d405-132">Depurar seu suplemento com o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="8d405-132">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="8d405-133">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="8d405-133">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="8d405-134">Depurar suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8d405-134">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)