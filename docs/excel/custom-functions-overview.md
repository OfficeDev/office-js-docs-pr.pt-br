---
ms.date: 09/26/2019
description: Criar funções personalizadas no Excel usando JavaScript.
title: Criar funções personalizadas no Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 252ff1badd935dda161f474bb7fefa8e782fd1c4
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814462"
---
# <a name="create-custom-functions-in-excel"></a>Criar funções personalizadas no Excel 

Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`. Este artigo descreve como criar as funções personalizadas no Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

A imagem animada a seguir mostra a sua pasta de trabalho solicitando uma função que você criou com o JavaScript ou o Typescript. Neste exemplo, a função personalizada `=MYFUNCTION.SPHEREVOLUME` calcula o volume de uma esfera.

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

O código a seguir define a função personalizada `=MYFUNCTION.SPHEREVOLUME`.

```js
/**
 * Returns the volume of a sphere. 
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!NOTE]
> A seção [Problemas conhecidos](#known-issues) neste artigo especifica as atuais limitações de funções personalizadas.

## <a name="how-a-custom-function-is-defined-in-code"></a>Como uma função personalizada é definida em código

Se você usar o [gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar funções personalizadas em um projeto do Excel, você encontrará que cria os arquivos que controlam as funções, o painel de tarefas e o suplemento geral. Vamos nos concentrar em arquivos que são importantes para funções personalizadas:

| File | Formato de arquivo | Descrição |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>ou<br/>**./src/functions/functions.ts** | JavaScript<br/>ou<br/>TypeScript | Contém o código que define funções personalizadas. |
| **./src/functions/functions.html** | HTML | Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas. |
| **./manifest.xml** | XML | Especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos JavaScript e HTML listados anteriormente nesta tabela. Também lista os locais de outros arquivos, que o suplemento pode fazer uso, como os arquivos do painel de tarefas e arquivos de comando. |

### <a name="script-file"></a>Arquivo de script

O arquivo de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contém o código que define funções e comentários que definem a função.

O código a seguir define a função personalizada `add`. Os comentários do código são usados para gerar um arquivo de metadados JSON que descreve a função personalizada ao Excel. O necessário `@customfunction` comentário é declarado primeiro, para indicar que se trata de uma função personalizada. Além disso, observe que dois parâmetros foram declarados, `first` e `second`, que é seguido por suas `description` propriedades. Por fim, uma `returns` descrição é fornecida. Para obter mais informações sobre quais comentários são necessários para sua função personalizada, confira [Criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}
```

Note que o arquivo **functions.html**, que governa o carregamento do tempo de execução das funções personalizadas, deve vincular-se à CDN atual para as funções personalizadas. Projetos preparados com a versão atual do gerador Yo Office referenciam a CDN correta. Se você estiver readaptando um projeto anterior de função personalizada de março de 2019 ou anteriormente, você precisará copiar no código abaixo para a página **functions.html**.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a>Arquivo de manifesto

O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto gerador que Yo Office cria) especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON.

A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas. Se estiver usando o gerador Yo Office, seus arquivos de funções personalizadas gerados conterão um arquivo de manifesto mais complexo, que você pode comparar neste [repositório do Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml).

> [!NOTE] 
> As URLs especificadas no arquivo de manifesto para as funções personalizadas JavaScript e JSON e arquivos HTML devem estar publicamente acessíveis e ter o mesmo subdomínio.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> Funções do Excel são anexadas ao namespace especificado no seu arquivo de manifesto XML. O namespace da função vem antes do nome da função e são separados por um ponto. Por exemplo, para acionar a função`ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque `CONTOSO` é o namespace e `ADD42` é o nome da função especificada no arquivo JSON. O namespace deve ser usado como identificador para o as sua empresa ou suplemento. Um namespace pode conter apenas caracteres alfanuméricos e períodos.

## <a name="coauthoring"></a>Coautoria

O Excel Online e no Windows conectado a uma assinatura do Office 365 permitem editar documentos em coautoria, e esse recurso funciona com funções personalizadas. Se a pasta de trabalho usa uma função personalizada, seu colega será solicitado a carregar o suplemento da função personalizada. Depois de carregarem o suplemento, a função personalizada compartilhará resultados por meio de coautoria.

Para saber mais sobre coautoria, confira o tópico [Sobre o recurso de coautoria no Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## <a name="known-issues"></a>Problemas conhecidos

Veja os problemas conhecidos no nosso [GitHub de funções do Excel personalizado repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="next-steps"></a>Próximas etapas

Quer experimentar funções personalizadas? Confira o simples [início rápido das funções personalizadas](../quickstarts/excel-custom-functions-quickstart.md) ou o mais detalhado [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md), caso ainda não tenha.

Outra maneira fácil de experimentar as funções personalizadas é usar o [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), que é um suplemento que permite com que você experimente as funções personalizadas diretamente no Excel. Você pode experimentar criar a sua própria função personalizada ou usar os exemplos disponíveis.

Pronto para ler mais sobre os recursos de funções personalizadas? Saiba mais sobre a visão geral da [arquitetura de funções personalizadas](custom-functions-architecture.md).

## <a name="see-also"></a>Confira também 
* [Requisitos de funções personalizadas](custom-functions-requirement-sets.md)
* [Diretrizes de nomenclatura](custom-functions-naming.md)
* [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](make-custom-functions-compatible-with-xll-udf.md)
