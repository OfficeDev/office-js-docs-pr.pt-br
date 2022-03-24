---
title: Especificar hosts do Office e requisitos de API
description: Saiba como especificar Office aplicativos e requisitos de API para que o seu complemento funcione conforme o esperado.
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1df753dfe3e5c517f49d597f9298744cf0c79f52
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744061"
---
# <a name="specify-office-applications-and-api-requirements"></a>Especificar requisitos da API e de aplicativos do Office

Seu Office Add-in pode depender de um aplicativo Office específico (também chamado de host Office) ou de membros específicos da API JavaScript do Office (office.js). Por exemplo, o suplemento pode:

- Executar em um único aplicativo do Office (por exemplo, Word ou Excel) ou diversos aplicativos.
- Use as OFFICE JAVAScript disponíveis apenas em algumas versões do Office. Por exemplo, a versão de compra única do Excel 2016 não oferece suporte a todas as APIs relacionadas Excel na biblioteca Office JavaScript.

Nessas situações, você precisa garantir que o seu add-in nunca seja instalado em aplicativos Office ou Office versões nas quais ele não possa ser executado.

Também há cenários em que você deseja controlar quais recursos do seu complemento ficam visíveis para os usuários com base no aplicativo Office e na versão Office. Dois exemplos são:

- Seu complemento tem recursos úteis no Word e no PowerPoint, como manipulação de texto, mas tem alguns recursos adicionais que só fazem sentido no PowerPoint, como recursos de gerenciamento de slides. Você precisa ocultar os recursos somente PowerPoint quando o complemento estiver em execução no Word.
- Seu add-in tem um recurso que requer um método de API JavaScript Office que é suportado em algumas versões de um aplicativo Office, como a assinatura Excel, mas não tem suporte em outras, como compra única Excel 2016. Mas o seu complemento tem outros recursos que exigem apenas Office de API JavaScript que são suportados  em Excel 2016. Nesse cenário, você precisa que o add-in seja instalado no Excel 2016, mas o recurso que exige o método sem suporte deve estar oculto dos usuários de Excel 2016.

Este artigo ajuda você a entender quais opções você deve escolher para garantir que seu suplemento funcione conforme o esperado e atinja o público mais amplo possível.

> [!NOTE]
> Para uma exibição de alto nível de onde os Office de Office atualmente são suportados, consulte Office disponibilidade de aplicativos cliente e plataforma para Office de [Office Desempate](../overview/office-add-in-availability.md).

> [!TIP]
> Muitas das tarefas descritas neste artigo são realizadas para você, no todo ou em parte, quando você cria seu projeto de complemento com uma ferramenta, como o gerador [Yeoman para complementos do Office](yeoman-generator-overview.md) ou um dos modelos de complemento do Office no Visual Studio. Nesses casos, interprete a tarefa como o que significa que você deve verificar se ela foi feita.

## <a name="use-the-latest-office-javascript-api-library"></a>Usar a biblioteca Office API JavaScript mais recente

O seu complemento deve carregar a versão mais atual da biblioteca de API JavaScript Office da rede de distribuição de conteúdo (CDN). Para fazer isso, certifique-se de que você tenha a seguinte `script` marca no primeiro arquivo HTML que seu complemento abrir. O uso de `/1/` na URL da CDN garante a referência à versão mais recente do Office.js.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>Especificar quais Office aplicativos podem hospedar seu complemento

Por padrão, um complemento pode ser instalado em todos os aplicativos Office suportados pelo tipo de complemento especificado (ou seja, Email, Painel de Tarefas ou Conteúdo). Por exemplo, um complemento do painel de tarefas é instalado por padrão no Access, Excel, OneNote, PowerPoint, Project e Word. 

Para garantir que o seu add-in seja instalado em um subconjunto de aplicativos Office, use os elementos [Hosts](../reference/manifest/hosts.md) e [Host](../reference/manifest/host.md) no manifesto.

Por exemplo, a seguinte declaração **hosts** e **host** especifica que o complemento pode instalar em qualquer versão do Excel, que inclui Excel na Web, Windows e iPad, mas não pode ser instalado em qualquer outro aplicativo Office.

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

O **elemento Hosts** pode conter um ou mais **elementos Host** . Deve haver um elemento **Host separado** para cada Office aplicativo no qual o add-in deve ser instalado. O `Name` atributo é necessário e pode ser definido como um dos seguintes valores.

| Nome          | Aplicativos cliente do Office                     | Tipos de complemento disponíveis |
|:--------------|:-----------------------------------------------|:-----------------------|
| Banco de dados      | Aplicativos Web do Access                                | Painel de tarefas              |
| Document      | Word na Web, Windows, Mac, iPad            | Painel de tarefas              |
| Mailbox       | Outlook na Web, Windows, Mac, Android, iOS | Correio                   |
| Notebook      | OneNote Online                             | Painel de tarefas, Conteúdo     |
| Presentation  | PowerPoint na Web, Windows, Mac, iPad      | Painel de tarefas, Conteúdo     |
| Project       | Project no Windows                             | Painel de tarefas              |
| Workbook      | Excel na Web, Windows, Mac, iPad           | Painel de tarefas, Conteúdo     |

> [!NOTE]
> Office aplicativos têm suporte em diferentes plataformas e são executados em desktops, navegadores da Web, tablets e dispositivos móveis. Normalmente, você não pode especificar qual plataforma pode ser usada para executar o seu complemento. Por exemplo, se você especificar `Workbook`, Excel na Web e Windows pode ser usado para executar o seu complemento. No entanto, se você especificar `Mailbox`, o seu add-in não será executado Outlook clientes móveis, a menos que você defina o [ponto de extensão móvel](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface).

> [!NOTE]
> Não é possível que um manifesto de complemento se aplique a mais de um tipo: Email, Painel de Tarefas ou Conteúdo. Isso significa que, se você quiser que o seu add-in seja instalado no Outlook e em um dos outros aplicativos Office, você deve criar dois complementos,  um com um manifesto de tipo Mail e outro com um painel de tarefas ou manifesto de tipo de conteúdo.

> [!IMPORTANT]
> Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint. Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>Especificar quais Office e plataformas podem hospedar seu complemento

Você não pode especificar explicitamente as versões e builds do Office ou as plataformas nas quais o seu complemento deve ser instalado, e você não gostaria, pois você teria que revisar seu manifesto sempre que o suporte para os recursos de complemento que seu complemento usa é estendido para uma nova versão ou plataforma. Em vez disso, especifique no manifesto as APIs que seu complemento precisa. Office impede que o add-Office in seja instalado em combinações de uma versão e uma plataforma que não suportam as APIs e garante que o complemento não apareça em **Meus Complementos**.

> [!IMPORTANT]
> Use apenas o manifesto base para especificar os membros da API que seu complemento deve ter de qualquer valor significativo. Se o seu complemento usa uma API para alguns recursos, mas tem outros recursos úteis que não exigem a API, você deve projetar o add-in para que ele seja instalado na plataforma e em combinações de versão Office que não suportam a API, mas fornece uma experiência diminuída nessas combinações. Para obter mais informações, consulte [Design for alternate experiences](#design-for-alternate-experiences).

### <a name="requirement-sets"></a>Conjuntos de requisitos

Para simplificar o processo de especificação das APIs que seu complemento precisa, Office a maioria das APIs em conjuntos *de requisitos*. As APIs no [Modelo de Objeto da API Comum](understanding-the-javascript-api-for-office.md#api-models) são agrupadas pelo recurso de desenvolvimento que eles suportam. Por exemplo, todas as APIs conectadas a vinculações de tabela estão no conjunto de requisitos chamado "TableBindings 1.1". As APIs nos modelos [de objeto específicos application](understanding-the-javascript-api-for-office.md#api-models) são agrupadas por quando elas foram lançadas para uso em complementos de produção.

Os conjuntos de requisitos são versionados. Por exemplo, as APIs que suportam [Caixas](../design/dialog-boxes.md) de Diálogo estão no conjunto de requisitos DialogApi 1.1. Quando apIs adicionais que habilitam mensagens de um painel de tarefas para uma caixa de diálogo foram lançadas, elas foram agrupadas em DialogApi 1.2, juntamente com todas as APIs em DialogApi 1.1. *Cada versão de um conjunto de requisitos é um superconjunto de todas as versões anteriores.*

O suporte ao conjunto de requisitos varia de acordo com Office aplicativo, a versão do aplicativo Office e a plataforma na qual ele está sendo executado. Por exemplo, o DialogApi 1.2 não é suportado em versões de compra única do Office antes do Office 2021, mas o DialogApi 1.1 é suportado em todas as versões de compra única de volta ao Office 2013. Você deseja que o seu add-in seja instalado em todas as combinações de plataforma e versão Office que suportam as APIs que ele usa, portanto, você sempre deve especificar no manifesto a  versão mínima de cada conjunto de requisitos que seu complemento exige. Detalhes sobre como fazer isso são posteriormente neste artigo.

> [!TIP]
> Para obter mais informações sobre o controle de versão do conjunto de requisitos, consulte [Office](office-versions-and-requirement-sets.md#office-requirement-sets-availability) disponibilidade de conjuntos de requisitos e para obter as listas completas de conjuntos de requisitos e informações sobre as APIs em cada uma delas, comece com Office [conjuntos](../reference/requirement-sets/office-add-in-requirement-sets.md) de requisitos de complemento. Os tópicos de referência para a maioria Office.js APIs também especificam o conjunto de requisitos ao qual pertencem (se algum).

> [!NOTE]
> Alguns conjuntos de requisitos também têm elementos de manifesto associados a eles. Consulte [Especificando requisitos em um elemento VersionOverrides](#specify-requirements-in-a-versionoverrides-element) para obter informações sobre quando esse fato é relevante para o design do seu complemento.

#### <a name="apis-not-in-a-requirement-set"></a>APIs que não estão em um conjunto de requisitos

Todas as APIs nos modelos específicos do aplicativo estão em conjuntos de requisitos, mas algumas delas no modelo de API comum não estão. Também há uma maneira de especificar uma dessas APIs sem definição no manifesto quando o seu complemento exigir uma. Os detalhes estão incluídos mais adiante neste artigo.

### <a name="requirements-element"></a>Elemento Requirements

Use o [elemento Requirements](../reference/manifest/requirements.md) e seus elementos filho [Conjuntos](../reference/manifest/sets.md) e Métodos para especificar os conjuntos de [requisitos mínimos](../reference/manifest/methods.md) ou membros da API que devem ser suportados pelo aplicativo Office para instalar o seu complemento. 

Se o aplicativo ou a plataforma Office não oferece suporte aos conjuntos de requisitos ou membros da API especificados no elemento **Requirements**, o complemento não será executado nesse aplicativo ou plataforma e não será exibido em **Meus Complementos**. .

> [!NOTE]
> O **elemento Requirements** é opcional para todos os complementos, exceto para Outlook de complementos. Quando o `xsi:type` atributo do elemento raiz `OfficeApp` `MailBox`for , deve haver um elemento **Requirements** que especifica a versão mínima do conjunto de requisitos mailbox que o complemento requer. Para obter mais informações, [consulte Outlook conjuntos de requisitos da API JavaScript](../reference/requirement-sets/outlook-api-requirement-sets.md).

O exemplo de código a seguir mostra como configurar um complemento que pode ser instalado em todos os aplicativos Office que suportam o seguinte:

-  `TableBindings` conjunto de requisitos, que tem uma versão mínima de "1.1".
-  `OOXML` conjunto de requisitos, que tem uma versão mínima de "1.1".
-  `Document.getSelectedDataAsync` método.

```XML
<OfficeApp ... >
  ...
  <Requirements>
     <Sets DefaultMinVersion="1.1">
        <Set Name="TableBindings" MinVersion="1.1"/>
        <Set Name="OOXML" MinVersion="1.1"/>
     </Sets>
     <Methods>
        <Method Name="Document.getSelectedDataAsync"/>
     </Methods>
  </Requirements>
    ...
</OfficeApp>
```
Observe o seguinte sobre este exemplo.

- O **elemento Requirements** contém os **elementos** filho **Conjuntos** e Métodos.
- O **elemento Sets** pode conter um ou mais **elementos Set** . `DefaultMinVersion` especifica o valor padrão `MinVersion` de todos os elementos **Set** filho.
- Um [elemento Set](../reference/manifest/set.md) especifica um conjunto de requisitos que o aplicativo Office deve dar suporte para tornar o add-in instaível. O `Name` atributo especifica o nome do conjunto de requisitos. Especifica `MinVersion` a versão mínima do conjunto de requisitos. `MinVersion` substitui o valor do atributo `DefaultMinVersion` nos Conjuntos **pai**.
- O **elemento Methods** pode conter um ou mais [elementos Method](../reference/manifest/method.md) . Você não pode usar o elemento **Methods** com suplementos do Outlook.
- O **elemento Method** especifica um método individual que o aplicativo Office deve dar suporte para tornar o add-in instaível. O `Name` atributo é necessário e especifica o nome do método qualificado com seu objeto pai.

## <a name="design-for-alternate-experiences"></a>Design para experiências alternativas

Os recursos de extensibilidade que a plataforma de Office add-in fornece podem ser divididos de forma útil em três tipos:

- Recursos de extensibilidade que estão disponíveis imediatamente após a instalação do complemento. Você pode usar esse tipo de recurso configurando um [elemento VersionOverrides](../reference/manifest/versionoverrides.md) no manifesto. Um exemplo desse tipo de recurso é [Comandos de Complemento](../design/add-in-commands.md), que são botões de faixa de opções personalizados e menus.
- Recursos de extensibilidade que estão disponíveis somente quando o add-in está em execução e que são implementados com Office.js APIs JavaScript; por exemplo, [Caixas de Diálogo](../design/dialog-boxes.md).
- Recursos de extensibilidade que estão disponíveis apenas no tempo de execução, mas são implementados com uma combinação de Office.js JavaScript e configuração em **um elemento VersionOverrides** . Exemplos disso são Excel [funções personalizadas](../excel/custom-functions-overview.md), [um único login](sso-in-office-add-ins.md) e [guias contextuais personalizadas](../design/contextual-tabs.md).

Se o seu add-in usa um recurso de extensibilidade específico para algumas de suas funcionalidades, mas tem outras funcionalidades úteis que não exigem o recurso de extensibilidade, você deve projetar o add-in para que ele seja instalado em combinações de versão de plataforma e Office que não suportam o recurso de extensibilidade. Ele pode fornecer uma experiência valiosa, embora diminuída, nessas combinações. 

Você implementa esse design de forma diferente, dependendo de como o recurso de extensibilidade é implementado: 

- Para recursos implementados inteiramente com JavaScript, consulte [Runtime checks for method and requirement set support](#runtime-checks-for-method-and-requirement-set-support).
- Para recursos que exigem que você configure **um elemento VersionOverrides** , consulte [Especificando requisitos em um elemento VersionOverrides](#specify-requirements-in-a-versionoverrides-element).

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>Verifica se há suporte ao método e ao conjunto de requisitos 

Você testa no tempo de execução para descobrir se a Office do usuário dá suporte a um conjunto de requisitos com o [método isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)). Passe o nome do conjunto de requisitos e a versão mínima como parâmetros. Se o conjunto de requisitos for suportado, `isSetSupported` **retornará true**. O código a seguir mostra um exemplo.

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
   // Code that uses API members from the WordApi 1.1 requirement set.
} else {
   // Provide diminished experience here. E.g., run alternate code when the user's Word is one-time purchase Word 2013 (which does not support WordApi 1.1).
}
```
Sobre este código, observe:

- O primeiro parâmetro é necessário. É uma cadeia de caracteres que representa o nome do conjunto de requisitos. Para saber mais sobre os conjuntos de requisitos disponíveis, confira [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md).
- O segundo parâmetro é opcional. É uma cadeia de caracteres que especifica a versão mínima do conjunto de requisitos que o aplicativo Office `if` deve suportar para que o código dentro da instrução seja executado (por exemplo, "**1.9**"). Se não for usada, a versão "1.1" será presumida.

> [!WARNING]
> Ao chamar o `isSetSupported` método, o valor do segundo parâmetro (se especificado) deve ser uma cadeia de caracteres e não um número. O analisador JavaScript não pode diferenciar entre valores numéricos como 1.1 e 1.10, enquanto ele pode para valores de cadeia de caracteres como "1.1" e "1.10".

A tabela a seguir mostra os nomes de conjunto de requisitos para os modelos de API específicos do aplicativo.

|Aplicativo do Office|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Caixa de correio|
|PowerPoint|PowerPointApi|
|Word|WordApi|

A seguir, um exemplo de uso do método com um dos conjuntos de requisitos de modelo de API comum.

```js
if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run alternate code when the user's Word doesn't support the CustomXmlParts requirement set.
}
```

> [!NOTE] 
> O `isSetSupported` método e os conjuntos de requisitos para esses aplicativos estão disponíveis no arquivo Office.js mais recente no CDN. Se você não usar Office.js do CDN, `isSetSupported` o seu complemento poderá gerar exceções se você estiver usando uma versão antiga da biblioteca na qual está indefinida. Para obter mais informações, [consulte Usar a biblioteca Office API JavaScript mais recente](#use-the-latest-office-javascript-api-library).

Quando o seu add-in depende de um método que não faz parte de um conjunto de requisitos, use a verificação de tempo de execução para determinar se o método é suportado pelo aplicativo Office, conforme mostrado no exemplo de código a seguir. Para obter uma lista completa dos métodos que não pertencem a um conjunto de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).

> [!NOTE]
> Recomendamos limitar o uso desse tipo de verificação no tempo de execução no código de seu suplemento.

O exemplo de código a seguir verifica se o aplicativo Office compatível `document.setSelectedDataAsync`com .

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>Especificar requisitos em um elemento VersionOverrides

O elemento [VersionOverrides](../reference/manifest/versionoverrides.md) foi adicionado ao esquema de manifesto principalmente, mas não exclusivamente, para dar suporte a recursos que devem estar disponíveis imediatamente após a instalação de um add-in, como comandos de complemento (botões de faixa de opções personalizados e menus). Office deve saber sobre esses recursos quando analisar o manifesto do complemento. 

Suponha que o seu complemento use um desses recursos, mas o complemento é valioso e deve ser instalado, mesmo em versões Office que não suportam o recurso. Neste cenário, identifique o recurso usando um elemento [Requirements](../reference/manifest/requirements.md) (e seus elementos Filho [Conjuntos](../reference/manifest/sets.md) e [](../reference/manifest/methods.md) Métodos) que você inclui como filho do próprio elemento **VersionOverrides** em vez de como filho do elemento base`OfficeApp`. O efeito de fazer isso é que o Office permitirá que o add-in seja instalado, mas o Office ignorará determinados dos elementos filho do **elemento VersionOverrides** em versões Office em que o recurso não é suportado.

Especificamente, os elementos filho dos **VersionOverrides** que substituem elementos no manifesto base, como um elemento **Hosts** , são ignorados e os elementos correspondentes do manifesto base são usados em vez disso. No entanto, pode haver elementos filho em **um VersionOverrides** que implementem recursos adicionais em vez de substituir as configurações no manifesto base. Dois exemplos são e `WebApplicationInfo` `EquivalentAddins`. Essas partes do **VersionOverrides** não serão  ignoradas, pressupondo que a plataforma e a versão do Office suportam o recurso correspondente.  

Para obter informações sobre os elementos descendentes do elemento **Requirements** , consulte [Elemento Requirements](#requirements-element) anteriormente neste artigo.

Apresentamos um exemplo a seguir.

```XML
<VersionOverrides ... >
   ...
   <Requirements>
      <Sets DefaultMinVersion="1.1">
         <Set Name="WordApi" MinVersion="1.2"/>
      </Sets>
   </Requirements>
   <Hosts>

      <!-- ALL MARKUP INSIDE THE HOSTS ELEMENT IS IGNORED WHEREVER WordApi 1.2 IS NOT SUPPORTED -->

      <Host xsi:type="Workbook">
         <!-- markup for custom add-in commands -->
      </Host>
   </Hosts>
</VersionOverrides>
```

> [!WARNING]
> Use um grande cuidado antes de usar um elemento **Requirements** em **um VersionOverrides**, porque em combinações de plataforma e versão que não suportam o *requisito, nenhum* dos comandos do add-in será instalado, mesmo aqueles que invocam a funcionalidade que não precisa do *requisito*. Considere, por exemplo, um complemento que tenha dois botões de faixa de opções personalizados. Uma delas chama Office APIs JavaScript disponíveis no conjunto de requisitos **ExcelApi 1.4** (e posterior). As outras CHAMADAS APIs que estão disponíveis apenas no **ExcelApi 1.9** (e posteriores). Se você colocar um requisito para **ExcelApi 1.9** no **VersionOverrides**, quando 1.9 não tiver suporte nenhum botão aparecerá na faixa de opções. Uma estratégia melhor nesse cenário seria usar a técnica descrita em Verificações de tempo de execução [para suporte ao método e ao conjunto de requisitos](#runtime-checks-for-method-and-requirement-set-support). O código invocado pelo segundo botão primeiro usa `isSetSupported` para verificar se há suporte do **ExcelApi 1.9**. Se não for suportado, o código dará ao usuário uma mensagem dizendo que esse recurso do complemento não está disponível em sua versão de Office. 

> [!TIP]
> Não faz sentido repetir um elemento **Requirement** em **um VersionOverrides** que já aparece no manifesto base. Se o requisito for especificado no manifesto base, o complemento não poderá instalar onde o requisito não é suportado para que Office nem mesmo analisar o elemento **VersionOverrides**. 

## <a name="see-also"></a>Confira também

- [Manifesto XML dos Suplementos do Office](add-in-manifests.md)
- [Conjuntos de requisitos de Suplemento do Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
