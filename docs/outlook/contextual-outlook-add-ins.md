---
title: Suplementos contextuais do Outlook
description: Inicie tarefas relacionadas a uma mensagem sem sair da mensagem para resultar em uma experiência de usuário mais fácil e mais sofisticada.
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 2f343f48f0c49de2b322cb737c5896df2f130ec9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747199"
---
# <a name="contextual-outlook-add-ins"></a>Suplementos contextuais do Outlook

Suplementos contextuais são suplementos do Outlook ativados com base no texto de um compromisso ou de uma mensagem. Usando suplementos contextuais, um usuário pode iniciar tarefas relacionadas a uma mensagem sem sair dela, o que resulta em uma experiência de usuário mais fácil e mais avançada.

A seguir estão exemplos de complementos contextuais.

- Escolher um endereço para abrir um mapa do local.
- Escolher uma cadeia de caracteres que abre um suplemento de sugestão de reunião.
- Escolher um número de telefone para adicionar aos seus contatos.


> [!NOTE]
> Atualmente, os suplementos contextuais não estão disponíveis no Outlook no Android e no iOS. Essa funcionalidade estará disponível no futuro.
>
> O suporte para esse recurso foi introduzido no conjunto de requisitos 1.6. Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="how-to-make-a-contextual-add-in"></a>Como fazer um suplemento contextual

O manifesto de um suplemento contextual deve conter um elemento [ExtensionPoint](../reference/manifest/extensionpoint.md#detectedentity) com um atributo `xsi:type` definido como `DetectedEntity`. No elemento **ExtensionPoint**, o suplemento especifica as entidades ou a expressão regular que podem ativá-lo. Se uma entidade for especificada, ela poderá ser qualquer uma das propriedades no objeto [Entities](/javascript/api/outlook/office.entities).

Dessa forma, o manifesto do suplemento precisa conter uma regra do tipo **ItemHasKnownEntity** ou **ItemHasRegularExpressionMatch**. O exemplo a seguir mostra como especificar que um complemento deve ser ativado em mensagens com uma entidade detectada que seja um número de telefone.

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
> Os tipos de entidade `EmailAddress` e `Url` não dão suporte ao realce, portanto, não podem ser usados para iniciar um suplemento contextual. No entanto, eles podem ser combinados em um tipo de regra `RuleCollection` como critérios de ativação adicionais.

## <a name="how-to-launch-a-contextual-add-in"></a>Como iniciar um suplemento contextual

O usuário inicia o suplemento contextual por meio de texto, tanto uma entidade conhecida quanto uma expressão regular do desenvolvedor. Normalmente, o usuário identifica um suplemento contextual porque a entidade está realçada. O exemplo a seguir mostra como o realce aparece em uma mensagem. Aqui, a entidade (um endereço) está na cor azul e sublinhada com uma linha pontilhada azul. Um usuário inicia o suplemento contextual clicando na entidade realçada. 

**Exemplo de texto com a entidade realçada (um endereço)**

![Mostra a entidade realçada em um email.](../images/outlook-detected-entity-highlight.png)
    
Quando há várias entidades ou suplementos contextuais em uma mensagem, existem algumas regras de interação do usuário:

- Se houver várias entidades, o usuário terá que clicar em uma entidade diferente para iniciar o suplemento.
- Se uma entidade ativar vários suplementos, cada suplemento abrirá uma nova guia. O usuário alterna entre guias para alternar entre os suplementos. Por exemplo, um nome e um endereço podem acionar um suplemento de telefone e um mapa.
- Se uma única cadeia de caracteres contiver várias entidades que ativam vários suplementos, toda a cadeia será realçada e um clique na cadeia de caracteres mostra todos os suplementos relevantes à cadeia em guias separadas. Por exemplo, uma cadeia de caracteres que descreve uma reunião proposta em um restaurante pode ativar o suplemento Reunião Sugerida e um suplemento de classificação de restaurantes.

## <a name="how-a-contextual-add-in-displays"></a>Como um suplemento contextual é exibido

Um suplemento contextual ativado aparece em um cartão, que é uma janela separada perto a entidade. O cartão normalmente aparecerá abaixo da entidade e centralizado o máximo possível em relação à entidade. Se não houver espaço suficiente embaixo da entidade, o cartão será colocado acima dela. A captura de tela a seguir mostra a entidade realçada e, abaixo dela, um suplemento (Bing Mapas) ativado em um cartão.

**Exemplo de um suplemento exibido em um cartão**

![Mostra um aplicativo contextual em um cartão.](../images/outlook-detected-entity-card.png)

Para fechar o cartão e o suplemento, o usuário deve clicar em algum lugar fora do cartão.

## <a name="current-contextual-add-ins"></a>Suplementos contextuais atuais

Os seguintes complementos contextuais são instalados por padrão para usuários com Outlook de complementos.

- Bing Mapas
- Reuniões sugeridas

## <a name="see-also"></a>Confira também

- [Suplemento do Outlook: número de ordem da Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (exemplo do suplemento contextual ativado com base em uma correspondência de expressão regular)
- [Escreva seu primeiro suplemento do Outlook](../quickstarts/outlook-quickstart.md)
- [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Objeto Entities](/javascript/api/outlook/office.entities)
