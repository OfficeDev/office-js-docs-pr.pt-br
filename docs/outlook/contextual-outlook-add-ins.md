---
title: Suplementos contextuais do Outlook
description: Inicie tarefas relacionadas a uma mensagem sem sair da mensagem para resultar em uma experiência de usuário mais fácil e mais sofisticada.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 73a13787dac7a6e74db6b919cc01a6dd33d29ab5
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467018"
---
# <a name="contextual-outlook-add-ins"></a>Suplementos contextuais do Outlook

Contextual add-ins are Outlook add-ins that activate based on text in a message or appointment. By using contextual add-ins, a user can initiate tasks related to a message without leaving the message itself, which results in an easier and richer user experience.

[!include[JSON manifest does not support contextual add-ins](../includes/json-manifest-outlook-contextual-not-supported.md)]

A seguir estão exemplos de suplementos contextuais.

- Escolher um endereço para abrir um mapa do local.
- Escolher uma cadeia de caracteres que abre um suplemento de sugestão de reunião.
- Escolher um número de telefone para adicionar aos seus contatos.


> [!NOTE]
> Atualmente, os suplementos contextuais não estão disponíveis no Outlook no Android e no iOS. Essa funcionalidade estará disponível no futuro.
>
> O suporte para esse recurso foi introduzido no conjunto de requisitos 1.6. Confira, [clientes e plataformas](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="how-to-make-a-contextual-add-in"></a>Como fazer um suplemento contextual

O manifesto de um suplemento contextual deve conter um elemento [ExtensionPoint](/javascript/api/manifest/extensionpoint#detectedentity) com um atributo `xsi:type` definido como `DetectedEntity`. Dentro do **\<ExtensionPoint\>** elemento, o suplemento especifica as entidades ou a expressão regular que pode ativá-lo. Se uma entidade for especificada, ela poderá ser qualquer uma das propriedades no objeto [Entities](/javascript/api/outlook/office.entities).

Dessa forma, o manifesto do suplemento precisa conter uma regra do tipo **ItemHasKnownEntity** ou **ItemHasRegularExpressionMatch**. O exemplo a seguir mostra como especificar que um suplemento deve ser ativado em mensagens com uma entidade detectada que seja um número de telefone.

```XML
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="contextLabel" />
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="detectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" Highlight="all" />
  </Rule>
</ExtensionPoint>
```

Depois que um suplemento contextual é associado a uma conta, ele inicia automaticamente quando o usuário clica em uma entidade ou expressão regular realçada. Para saber mais sobre expressões regulares para Suplementos do Outlook, confira [Usar regras de ativação de expressões regulares para mostrar um Suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).

Há várias restrições em suplementos contextuais:

- Um suplemento contextual só pode existir em suplementos de leitura (não de redação).
- Você não pode especificar a cor da entidade realçada.
- Uma entidade que não estiver realçada não iniciará um suplemento contextual em um cartão.

Como uma entidade ou expressão regular que não estiver realçada não iniciará o suplemento contextual, os suplementos devem conter pelo menos um elemento `Rule` com o atributo `Highlight` definido como `all`.

> [!NOTE]
> The `EmailAddress` and `Url` entity types do not support highlighting, so they cannot be used to launch a contextual add-in. They can however be combined in a `RuleCollection` rule type as an additional activation criteria.

## <a name="how-to-launch-a-contextual-add-in"></a>Como iniciar um suplemento contextual

A user launches a contextual add-in through text, either a known entity or a developer's regular expression. Typically, a user identifies a contextual add-in because the entity is highlighted. The following example shows how highlighting appears in a message. Here the entity (an address) is colored blue and underlined with a dotted blue line. A user launches the contextual add-in by clicking the highlighted entity. 

**Exemplo de texto com a entidade realçada (um endereço)**

![Mostra a entidade realçada em um email.](../images/outlook-detected-entity-highlight.png)
    
Quando há várias entidades ou suplementos contextuais em uma mensagem, existem algumas regras de interação do usuário:

- Se houver várias entidades, o usuário terá que clicar em uma entidade diferente para iniciar o suplemento.
- Se uma entidade ativar vários suplementos, cada suplemento abrirá uma nova guia. O usuário alterna entre guias para alternar entre os suplementos. Por exemplo, um nome e um endereço podem acionar um suplemento de telefone e um mapa.
- If a single string contains multiple entities that activate multiple add-ins, the entire string is highlighted, and clicking the string shows all add-ins relevant to the string on separate tabs. For example, a string that describes a proposed meeting at a restaurant might activate the Suggested Meeting add-in and a restaurant rating add-in.

## <a name="how-a-contextual-add-in-displays"></a>Como um suplemento contextual é exibido

An activated contextual add-in appears in a card, which is a separate window near the entity. The card will normally appear below the entity and centered with respect to the entity as much as possible. If there is not enough room below the entity, the card is placed above it. The following screenshot shows the highlighted entity, and below it, an activated add-in (Bing Maps) in a card.

**Exemplo de um suplemento exibido em um cartão**

![Mostra um aplicativo contextual em um cartão.](../images/outlook-detected-entity-card.png)

Para fechar o cartão e o suplemento, o usuário deve clicar em algum lugar fora do cartão.

## <a name="current-contextual-add-ins"></a>Suplementos contextuais atuais

Os suplementos contextuais a seguir são instalados por padrão para usuários com suplementos do Outlook.

- Bing Mapas
- Reuniões sugeridas

## <a name="see-also"></a>Confira também

- [Suplemento do Outlook: número de ordem da Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (exemplo do suplemento contextual ativado com base em uma correspondência de expressão regular)
- [Escreva seu primeiro suplemento do Outlook](../quickstarts/outlook-quickstart.md)
- [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Objeto Entities](/javascript/api/outlook/office.entities)
