---
title: Visão geral dos suplementos do Word
description: Conheça as noções básicas de suplementos do Word
ms.date: 11/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: b27de57c821fe8278ac8a6e097aeb8e5be0cdb7b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717324"
---
# <a name="word-add-ins-overview"></a>Visão geral dos suplementos do Word

Deseja criar uma solução que amplie a funcionalidade do Word? Por exemplo, uma que envolva montagem automatizada de documentos? Ou uma solução que vincule e acesse dados em um documento do Word a partir de outras fontes de dados? Você pode usar a plataforma de Suplementos do Office, que inclui a API JavaScript do Word e a API JavaScript do Office, para estender os clientes executando o Word na área de trabalho do Windows, no Mac ou na nuvem.

Os suplementos do Word são uma das várias opções de desenvolvimento disponíveis na [plataforma de suplementos do Office](../overview/office-add-ins.md). Você pode usar comandos de suplemento para estender a interface do usuário do Word e iniciar os painéis de tarefas que executam JavaScript que interage com o conteúdo em um documento do Word. Qualquer código que você pode executar em um navegador, pode ser executado em um suplemento do Word. Suplementos que interagem com conteúdo em um documento do Word criam solicitações para agir em objetos do Word e sincronizar o estado do objeto. 

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento no AppSource depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade do suplemento do Office](../overview/office-add-in-availability.md)).

A figura a seguir mostra um exemplo de um suplemento do Word que é executado em um painel de tarefas.

*Figura 1. Suplemento em execução em um painel de tarefas no Word*

![Suplemento em execução em um painel de tarefas no Word](../images/word-add-in-show-host-client.png)

O suplemento do Word (1) pode enviar solicitações para o documento do Word (2) e usar o JavaScript para acessar o objeto parágrafo e atualizar, excluir ou mover o parágrafo. Por exemplo, o código a seguir mostra como acrescentar uma nova sentença a esse parágrafo.

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

É possível usar qualquer tecnologia de servidor Web para hospedar o suplemento do Word, como ASP.NET, NodeJS ou Python. Use a estrutura de cliente de sua preferência (Ember, Backbone, Angular, React) ou use o VanillaJS para desenvolver a solução. É possível usar serviços como o Azure para [autenticar](../develop/overview-authn-authz.md) e hospedar seu aplicativo.

As APIs JavaScript do Word proporcionam ao seu aplicativo o acesso aos objetos e metadados encontrado em um documento do Word. Você pode usar essas APIs para criar suplementos que têm como objetivo:

* Word 2013 ou posterior no Windows
* Word Online
* Word 2016 ou posterior no Windows
* Word no iPad

Redija seu suplemento uma vez e ele será executado em todas as versões do Word em várias plataformas. Para obter detalhes, consulte [Disponibilidade de Suplementos do Office em hosts e plataformas](../overview/office-add-in-availability.md).

## <a name="javascript-apis-for-word"></a>APIs JavaScript para Word

Você pode usar dois conjuntos de APIs JavaScript para interagir com metadados e objetos em um documento do Word. O primeiro é a [API comuns](/javascript/api/office), que foi introduzido no Office 2013. Muitos dos objetos na API comuns podem ser usados em suplementos hospedados por dois ou mais clientes do Office.  Essa API usa retornos de chamadas de maneira ampla.

O segundo é a [API JavaScript do Word](/javascript/api/word). Este é um modelo de objeto fortemente tipado que você pode usar para criar suplementos do Word que se destinam ao Word 2016 para Mac e Windows. Este modelo de objeto usa promessas e fornece acesso a objetos específicos do Word como [corpo](/javascript/api/word/word.body), [controles de conteúdo](/javascript/api/word/word.contentcontrol), [imagens embutidas](/javascript/api/word/word.inlinepicture) e [parágrafos](/javascript/api/word/word.paragraph). A API JavaScript do Word inclui definições do TypeScript e arquivos vsdoc para que você possa obter dicas de código em seu IDE.

Atualmente, todos os clientes do Word oferecem suporte à API JavaScript do Office compartilhada, e a maioria dos clientes oferece suporte à API JavaScript do Word. Para obter detalhes sobre clientes com suporte, consulte [Disponibilidade do host e da plataforma dos Suplementos do Office](../overview/office-add-in-availability.md).

Recomendamos que você comece com a API JavaScript do Word porque o modelo de objeto é mais fácil de usar. Use a API JavaScript do Word se precisar:

* Acessar os objetos em um documento do Word.

Use a API JavaScript do Office compartilhada quando precisar:

* Direcionar o Word 2013.
* Executar ações iniciais do aplicativo.
* Verificar o conjunto requisitos com suporte.
* Acessar metadados, configurações e informações do ambiente para o documento.
* Vincular a seções em um documento e capturar eventos.
* Usar partes XML personalizadas.
* Abrir uma caixa de diálogo.

## <a name="next-steps"></a>Próximas etapas

Pronto para criar seu primeiro suplemento do Word? Confira [Criar seu primeiro suplemento do Word](word-add-ins.md). Use o [manifesto de suplemento](../develop/add-in-manifests.md) para descrever onde seu suplemento está hospedado e como ele é exibido, bem como para definir permissões e outras informações.

Para saber mais sobre como projetar um suplemento do Word de classe internacional que cria uma ótima experiência para seus usuários, consulte [Diretrizes de design](../design/add-in-design.md) e [Práticas recomendadas](../concepts/add-in-development-best-practices.md).

Depois de desenvolver seu suplemento, é possível [publicá-lo](../publish/publish.md) em um compartilhamento de rede, um catálogo de aplicativos ou no AppSource.

## <a name="see-also"></a>Confira também

* [Criando Suplementos do Office ](../overview/office-add-ins-fundamentals.md)
* [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
* [Referências da API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md)
