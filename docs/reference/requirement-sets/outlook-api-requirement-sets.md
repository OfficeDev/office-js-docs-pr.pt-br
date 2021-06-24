---
title: Conjuntos de requisitos de API JavaScript do Outlook
description: Saiba mais sobre os conjuntos de requisitos da API JavaScript do Outlook.
ms.date: 05/17/2021
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 49cfcfee075ba01f077162cef415ed58211b95f6
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077208"
---
# <a name="outlook-javascript-api-requirement-sets"></a>Conjuntos de requisitos de API JavaScript do Outlook

Os Suplementos do Outlook declaram quais versões de API exigem usando o elemento Requisitos em seu manifesto. Os suplementos do Outlook sempre incluem um elemento Conjunto com um atributo  definido como  e um atributo  definido como o conjunto de requisitos mínimo de API compatível com os cenários do suplemento.

Por exemplo, o trecho de código de manifesto a seguir indica um conjunto mínimo de requisitos de 1.1.

```xml
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Todas as APIs do Outlook pertencem ao [conjunto de requisitos](../../develop/specify-office-hosts-and-api-requirements.md)`Mailbox`. O conjunto de requisitos `Mailbox` tem versões, e cada novo conjunto de APIs que lançamos pertence a uma versão superior. Nem todos os clientes do Outlook serão compatíveis com o conjunto mais recente de APIs, mas se um cliente do Outlook declarar suporte a um conjunto de requisitos, será compatível com todas as APIs nesse conjunto (verifique a documentação de uma API ou recurso específicos para qualquer possível exceção).

A especificação de uma versão mínima de conjunto de requisitos controla em quais clientes do Outlook o suplemento aparecerá. Se um cliente não oferece suporte para o conjunto de requisitos mínimos, ele não carrega o suplemento. Por exemplo, se for especificada a versão 1.3 do conjunto de requisitos, significa que o suplemento não aparecerá nos clientes do Outlook incompatíveis com a versão 1.3.

> [!NOTE]
> Para usar APIs em qualquer um dos conjuntos de requisitos numerados, faça referência à biblioteca **production** no CDN (https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> Para obter informações sobre o uso de APIs de visualização, confira a seção [Usando APIs de visualização ](#using-preview-apis), mais adiante neste artigo.

## <a name="using-apis-from-later-requirement-sets"></a>Usar APIs de conjuntos de requisitos posteriores

Definir um conjunto de requisitos não limita as APIs disponíveis que o suplemento pode usar. Por exemplo, se o suplemento especifica o conjunto de requisitos "Caixa de correio 1.1", mas está sendo executado em um cliente Outlook que oferece suporte a "Caixa de correio 1.3", o suplemento pode usar APIs do conjunto de requisitos "Caixa de correio 1.3".

Para usar uma API mais recente, os desenvolvedores podem verificar se um determinado aplicativo oferece suporte ao conjunto de requisitos fazendo o seguinte:

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

Como alternativa, os desenvolvedores podem verificar a existência de uma API mais recente usando a técnica JavaScript padrão.

```js
if (item.somePropertyOrFunction !== undefined) {
  // Use item.somePropertyOrFunction.
  item.somePropertyOrFunction;
}
```

Nenhuma verificação desse tipo é necessária para qualquer API que esteja presente na versão do conjunto de requisitos especificada no manifesto.

## <a name="choosing-a-minimum-requirement-set"></a>Escolher um conjunto de requisitos mínimos

Os desenvolvedores devem usar o conjunto de requisitos mínimos que contém o conjunto essencial de APIs para seu cenário, sem o qual o suplemento não funcionará.

## <a name="requirement-sets-supported-by-exchange-servers-and-outlook-clients"></a>Conjuntos de requisitos suportados pelos Exchange Servers e clientes do Outlook

Nesta seção, vemos a gama de conjuntos de requisitos com suporte do Exchange Server e clientes do Outlook. Para obter detalhes sobre os requisitos de cliente e servidor para executar suplementos do Outlook, confira [requisitos dos suplementos do Outlook](../../outlook/add-in-requirements.md).

> [!IMPORTANT]
> Se o seu Exchange Server de destino e o cliente do Outlook oferecem suporte a conjuntos de requisitos diferentes, então você estará restrito ao intervalo menor de conjunto de requisitos. Por exemplo, se um suplemento estiver sendo executado no Outlook 2016 para Mac (conjunto de requisitos mais alto: 1.6) em relação ao Exchange 2013 (conjunto de requisitos mais alto: 1.1), seu suplemento estará limitado ao conjunto de requisitos 1.1.

### <a name="exchange-server-support"></a>Suporte do Exchange Server

Os clientes a seguir oferecem suporte aos suplementos do Outlook.

| Produto | Versão Principal do Exchange | Conjuntos de requisitos de API com suporte |
|---|---|---|
| Exchange Online | Versão mais recente | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](../objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)\* |
| Exchange local | 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md) |

> [!NOTE]
> \* Para exigir o conjunto 1.3 da API de Identidade no código do suplemento, verifique se ele tem suporte ligando para `isSetSupported('IdentityAPI', '1.3')`. Não há suporte para declará-lo no manifesto do seu suplemento. Você também pode determinar se a API tem suporte, verificando se ela não é `undefined`. Para mais detalhes, confira [Usar APIs de conjuntos de requisitos posteriores](#using-apis-from-later-requirement-sets).

### <a name="outlook-client-support"></a>Suporte a cliente Outlook

Os suplementos são compatíveis com o Outlook nas seguintes plataformas.

| Plataforma | Versão Principal do Office/Outlook | Conjuntos de requisitos de API com suporte |
|---|---|---|
| Windows | Assinatura do Microsoft 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup>, [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)<sup>1</sup>, [1.10](../objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<sup>1</sup><br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| Compra única de 2019 (varejo) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup> |
|| Compra única de 2019 (licenciada por volume) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md) |
|| compra avulsa 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
|| compra avulsa 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)<sup>3</sup>, [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
| Mac | Interface de usuário atual<br>(conectada a uma assinatura do Microsoft 365) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| Nova interface de usuário (pré-visualização)<sup>4</sup><br>(conectado a uma assinatura do Microsoft 365) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| compra avulsa 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| compra avulsa 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
| iOS | Assinatura do Microsoft 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Android | Assinatura do Microsoft 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Navegador da Web | interface do usuário moderna do Outlook quando conectado ao<br>Exchange Online: assinatura do Microsoft 365, Outlook.com | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](../objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| interface do usuário clássica do Outlook quando conectado ao<br>Exchange local | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |

> [!NOTE]
> <sup>1</sup> Suporte para **1.8** do Outlook no Windows com uma assinatura Microsoft 365 ou uma compra única no varejo está disponível a partir da versão 1910 (compilação 12130.20272). O suporte para **1.9** do Outlook no Windows com uma assinatura Microsoft 365 está disponível a partir da versão 2008 (compilação 13127.20296). O suporte para **1.10** do Outlook no Windows com uma assinatura Microsoft 365 está disponível a partir da versão 2104 (compilação 13929.20296). Para obter mais detalhes em relação à sua versão, consulte a página do histórico de atualizações do [Office 2019](/officeupdates/update-history-office-2019) ou [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) e de como [encontrar a versão do cliente do Office e atualizar o canal](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).
>
> <sup>2</sup> Para exigir o conjunto 1.3 da API de Identidade no código do suplemento, verifique se ele tem suporte ligando para `isSetSupported('IdentityAPI', '1.3')`. Não há suporte para declará-lo no manifesto do seu suplemento. Você também pode determinar se a API tem suporte, verificando se ela não é `undefined`. Para mais detalhes, confira [Usar APIs de conjuntos de requisitos posteriores](#using-apis-from-later-requirement-sets).
>
> <sup>3</sup> O suporte de 1.3 no Outlook 2013 foi adicionado como parte da atualização de [8 de dezembro de 2015 do Outlook 2013 (KB3114349).](https://support.microsoft.com/kb/3114349). O suporte para a versão 1.4 no Outlook 2013 foi adicionado como parte da [atualização para Outlook 2013 de 13 de setembro de 2016 (KB3118280)](https://support.microsoft.com/help/3118280). O suporte para a versão 1.4 no Outlook 2016 (compra única) foi adicionado como parte da [atualização para o Office 2016 de 3 de julho de 2018 (KB4022223)](https://support.microsoft.com/help/4022223).
>
> <sup>4</sup> O suporte para a nova Interface do Usuário do Mac (visualização) está disponível no Outlook versão 16.38.506. Para mais informações, consulte a seção [Suporte de Suplemento no Outlook na nova Interface do Usuário do Mac](../../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview).
>
> <sup>5</sup> No momento, existem considerações adicionais ao se projetar e implementar suplementos para clientes móveis. Por exemplo, o único modo suportado é o Message Read. Para obter mais detalhes, consulte [considerações de código ao adicionar suporte aos comandos de suplemento do Outlook Mobile](../../outlook/add-mobile-support.md#code-considerations).

> [!TIP]
> É possível distinguir o Outlook clássico do moderno no navegador da Web, verificando sua barra de ferramentas da caixa de correio.
>
> **moderno**
>
> ![Captura de tela parcial da barra de ferramentas moderna do Outlook.](../../images/outlook-on-the-web-new-toolbar.png)
>
> **clássico**
>
> ![Captura de tela parcial da barra de ferramentas clássica do Outlook.](../../images/outlook-on-the-web-classic-toolbar.png)

## <a name="using-preview-apis"></a>Usando APIs de visualização

As novas APIs do JavaScript para Outlook são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários. Para fornecer feedback sobre uma API de visualização, use o mecanismo de feedback no final da página da Web em que a API está documentada.

> [!NOTE]
> As APIs de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção.

Para saber mais detalhes sobre as APIs de visualização, confira o artigo sobre o [conjunto de requisitos da API de visualização do Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md).
