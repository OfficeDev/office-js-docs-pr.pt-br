---
title: Visão geral dos suplementos do Word
description: ''
ms.date: 09/24/2018
ms.openlocfilehash: 5cfae87c44f2a3004e4cd755614d15261b43945f
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505941"
---
# <a name="word-add-ins-overview"></a>Visão geral dos suplementos do Word

Você deseja criar uma solução que estenda a funcionalidade do Word? Por exemplo, uma solução para montagem automatizada de documentos? Ou uma solução que associe e acesse dados em um documento Word a partir de outras fontes de dados? Você pode usar a plataforma de suplementos do Office, que inclui a API JavaScript do Word e a API JavaScript para Office, para ampliar os clientes do Word no Windows desktop, em Mac ou na nuvem.

Os suplementos do Word são uma das várias opções de desenvolvimento disponíveis na [plataforma de suplementos do Office](../overview/office-add-ins.md). Você pode usar comandos de suplemento para estender a interface do usuário do Word e iniciar os painéis de tarefas que executam o JavaScript que interage com o conteúdo em um documento do Word. Todos os códigos que podem ser executados em um navegador, também podem ser executados em um suplemento do Word. Os suplementos que interagem com o conteúdo em um documento do Word criam solicitações para agir em objetos do Word e sincronizar o estado do objeto. 

> [!NOTE]
> Se pretende [publicar](../publish/publish.md) o suplemento no AppSource depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade do suplemento do Office](../overview/office-add-in-availability.md)).

A figura a seguir mostra um exemplo de um suplemento do Word que é executado em um painel de tarefas.

*Figura 1. Suplemento em execução em um painel de tarefas no Word*

![Suplemento em execução em um painel de tarefas no Word](../images/word-add-in-show-host-client.png)

O suplemento do Word (1) pode enviar solicitações para o documento do Word (2) e usar o JavaScript para acessar o objeto paragraph e atualizar, excluir ou mover o parágrafo. Por exemplo, o código a seguir mostra como acrescentar uma nova sentença a esse parágrafo.

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

É possível usar qualquer tecnologia de servidor web para hospedar o suplemento do Word, como ASP.NET, NodeJS ou Python. Use a estrutura de cliente de sua preferência (Ember, Backbone, Angular, React) ou use o VanillaJS para desenvolver a solução. Pode usar serviços como o Azure para [autenticar](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) e hospedar o aplicativo.

As APIs JavaScript do Word fornecem ao seu aplicativo acesso aos objetos e metadados encontrados em um documento do Word. Você pode usar essas APIs para criar suplementos que segmentam:

* Word 2013 ou posterior para Windows
* Word Online
* Word 2016 ou posterior para Mac
* Word para iOS

Escreva seu suplemento uma vez e ele será executado em todas as versões do Word em várias plataformas. Para obter detalhes, consulte [Disponibilidade de suplementos do Office em hosts e plataformas](../overview/office-add-in-availability.md).

## <a name="javascript-apis-for-word"></a>APIs JavaScript para Word

Você pode usar dois APIs JavaScript para interagir com metadados e objetos em um documento do Word. A primeira é a [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js?product=word), que foi introduzida no Office 2013. É uma API compartilhada – muitos dos objetos podem ser usados em suplementos hospedados por dois ou mais clientes do Office. Essa API usa retornos de chamadas de forma ampla.

A segunda é a [API JavaScript do Word](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js). Essa é um modelo de objeto fortemente tipado que você pode usar para criar suplementos do Word que se destinam ao Word 2016 para Mac e Windows. O modelo de objeto usa promessas e fornece acesso a objetos específicos do Word como [body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js), [content controls](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js), [inline pictures](https://docs.microsoft.com/javascript/api/word/word.inlinepicture?view=office-js), e [paragraphs](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js). A API JavaScript do Word inclui definições de TypeScript e arquivos vsdoc para que você possa obter dicas de código em seu IDE.

Atualmente, todos os clientes do Word oferecem suporte à API JavaScript para Office compartilhada, e a maioria dos clientes oferece suporte à API JavaScript do Word. Para obter detalhes sobre clientes com suporte, consulte a [documentação de referência da API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js?product=word).

Recomendamos que você comece com a API JavaScript do Word porque o modelo de objeto é mais fácil de usar. Use a API JavaScript do Word se precisa:

* Acessar os objetos em um documento do Word.

Use a API JavaScript para Office compartilhada se precisa:

* Direcionar o Word 2013.
* Executar ações iniciais para o aplicativo.
* Verificar o conjunto de requisitos com suporte.
* Acessar metadados, configurações e informações do ambiente para o documento.
* Vincular a seções em um documento e capturar eventos.
* Usar partes XML personalizadas.
* Abrir uma caixa de diálogo.

## <a name="next-steps"></a>Próximas etapas

Pronto para criar seu primeiro suplemento do Word? Confira [Compilar seu primeiro suplemento do Word](word-add-ins.md). Também pode acessar nossa [Experiência de introdução](https://docs.microsoft.com/office/dev/add-ins/?product=Word) interativa. Use um [manifesto do suplemento](../develop/add-in-manifests.md) para descrever onde seu suplemento está hospedado e como é exibido, além de definir permissões e outras informações.

Para saber mais sobre como criar um suplemento do Word de classe mundial que ofereça uma ótima experiência para seus usuários, consulte [Diretrizes de design](../design/add-in-design.md) e [Práticas recomendadas](../concepts/add-in-development-best-practices.md).

Depois de desenvolver seu suplemento, é possível [publicá-lo](../publish/publish.md) em um compartilhamento de rede, um catálogo de aplicativos ou no AppSource.

## <a name="whats-coming-up-for-word-add-ins"></a>O que está surgindo para os suplementos do Word?

À medida que projetamos e desenvolvemos novas APIs para os suplementos do Word, as disponibilizamos na nossa página [Especificações abertas das APIs](https://docs.microsoft.com/office/dev/add-ins/reference/openspec?view=office-js) para você deixar seus comentários. Descubra os novos recursos que estão no pipeline para as APIs JavaScript do Word e forneça comentários sobre nossas especificações de design.

## <a name="see-also"></a>Confira também

* [Visão geral da plataforma de suplementos do Office](../overview/office-add-ins.md)
* [Referência da API JavaScript do Word](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js)

