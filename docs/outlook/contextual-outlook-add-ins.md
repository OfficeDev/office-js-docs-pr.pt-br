---
title: Suplementos contextuais do Outlook
description: Inicie tarefas relacionadas a uma mensagem sem sair da mensagem para resultar em uma experiência de usuário mais fácil e mais sofisticada.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: c9a01e05fa5bb0a0932da50b096fa2cb71cf3b34
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076774"
---
# <a name="contextual-outlook-add-ins"></a><span data-ttu-id="a03f0-103">Suplementos contextuais do Outlook</span><span class="sxs-lookup"><span data-stu-id="a03f0-103">Contextual Outlook add-ins</span></span>

<span data-ttu-id="a03f0-p101">Suplementos contextuais são suplementos do Outlook ativados com base no texto de um compromisso ou de uma mensagem. Usando suplementos contextuais, um usuário pode iniciar tarefas relacionadas a uma mensagem sem sair dela, o que resulta em uma experiência de usuário mais fácil e mais avançada.</span><span class="sxs-lookup"><span data-stu-id="a03f0-p101">Contextual add-ins are Outlook add-ins that activate based on text in a message or appointment. By using contextual add-ins, a user can initiate tasks related to a message without leaving the message itself, which results in an easier and richer user experience.</span></span>

<span data-ttu-id="a03f0-106">A seguir apresentamos exemplos de suplementos contextuais:</span><span class="sxs-lookup"><span data-stu-id="a03f0-106">The following are examples of contextual add-ins:</span></span>

- <span data-ttu-id="a03f0-107">Escolher um endereço para abrir um mapa do local.</span><span class="sxs-lookup"><span data-stu-id="a03f0-107">Choosing an address to open a map of the location.</span></span>
- <span data-ttu-id="a03f0-108">Escolher uma cadeia de caracteres que abre um suplemento de sugestão de reunião.</span><span class="sxs-lookup"><span data-stu-id="a03f0-108">Choosing a string that opens a meeting suggestion add-in.</span></span>
- <span data-ttu-id="a03f0-109">Escolher um número de telefone para adicionar aos seus contatos.</span><span class="sxs-lookup"><span data-stu-id="a03f0-109">Choosing a phone number to add to your contacts.</span></span>


> [!NOTE]
> <span data-ttu-id="a03f0-110">Atualmente, os suplementos contextuais não estão disponíveis no Outlook no Android e no iOS.</span><span class="sxs-lookup"><span data-stu-id="a03f0-110">Contextual add-ins are not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="a03f0-111">Essa funcionalidade estará disponível no futuro.</span><span class="sxs-lookup"><span data-stu-id="a03f0-111">This functionality will be made available in the future.</span></span>
>
> <span data-ttu-id="a03f0-112">O suporte para esse recurso foi introduzido no conjunto de requisitos 1.6.</span><span class="sxs-lookup"><span data-stu-id="a03f0-112">Support for this feature was introduced in requirement set 1.6.</span></span> <span data-ttu-id="a03f0-113">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="a03f0-113">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="how-to-make-a-contextual-add-in"></a><span data-ttu-id="a03f0-114">Como fazer um suplemento contextual</span><span class="sxs-lookup"><span data-stu-id="a03f0-114">How to make a contextual add-in</span></span>

<span data-ttu-id="a03f0-115">O manifesto de um suplemento contextual deve conter um elemento [ExtensionPoint](../reference/manifest/extensionpoint.md#detectedentity) com um atributo `xsi:type` definido como `DetectedEntity`.</span><span class="sxs-lookup"><span data-stu-id="a03f0-115">A contextual add-in's manifest must include an [ExtensionPoint](../reference/manifest/extensionpoint.md#detectedentity) element with an `xsi:type` attribute set to `DetectedEntity`.</span></span> <span data-ttu-id="a03f0-116">No elemento **ExtensionPoint**, o suplemento especifica as entidades ou a expressão regular que podem ativá-lo.</span><span class="sxs-lookup"><span data-stu-id="a03f0-116">Within the **ExtensionPoint** element, the add-in specifies the entities or regular expression that can activate it.</span></span> <span data-ttu-id="a03f0-117">Se uma entidade for especificada, ela poderá ser qualquer uma das propriedades no objeto [Entities](/javascript/api/outlook/office.entities).</span><span class="sxs-lookup"><span data-stu-id="a03f0-117">If an entity is specified, the entity can be any of the properties in the [Entities](/javascript/api/outlook/office.entities) object.</span></span>

<span data-ttu-id="a03f0-118">Dessa forma, o manifesto do suplemento precisa conter uma regra do tipo **ItemHasKnownEntity** ou **ItemHasRegularExpressionMatch**.</span><span class="sxs-lookup"><span data-stu-id="a03f0-118">Thus, the add-in manifest must contain a rule of type **ItemHasKnownEntity** or **ItemHasRegularExpressionMatch**.</span></span> <span data-ttu-id="a03f0-119">O exemplo a seguir mostra como especificar que um suplemento deve se ativar em mensagens com uma entidade detectada que é um número de telefone:</span><span class="sxs-lookup"><span data-stu-id="a03f0-119">The following example shows how to specify that an add-in should activate on messages with a detected entity that is a phone number:</span></span>

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

<span data-ttu-id="a03f0-120">Depois que um suplemento contextual é associado a uma conta, ele inicia automaticamente quando o usuário clica em uma entidade ou expressão regular realçada.</span><span class="sxs-lookup"><span data-stu-id="a03f0-120">After a contextual add-in is associated with an account, it will automatically start when the user clicks a highlighted entity or regular expression.</span></span> <span data-ttu-id="a03f0-121">Para saber mais sobre expressões regulares para Suplementos do Outlook, confira [Usar regras de ativação de expressões regulares para mostrar um Suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="a03f0-121">For more information about regular expressions for Outlook add-ins, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>

<span data-ttu-id="a03f0-122">Há várias restrições em suplementos contextuais:</span><span class="sxs-lookup"><span data-stu-id="a03f0-122">There are several restrictions on contextual add-ins:</span></span>

- <span data-ttu-id="a03f0-123">Um suplemento contextual só pode existir em suplementos de leitura (não de redação).</span><span class="sxs-lookup"><span data-stu-id="a03f0-123">A contextual add-in can only exist in read add-ins (not compose add-ins).</span></span>
- <span data-ttu-id="a03f0-124">Você não pode especificar a cor da entidade realçada.</span><span class="sxs-lookup"><span data-stu-id="a03f0-124">You cannot specify the color of the highlighted entity.</span></span>
- <span data-ttu-id="a03f0-125">Uma entidade que não estiver realçada não iniciará um suplemento contextual em um cartão.</span><span class="sxs-lookup"><span data-stu-id="a03f0-125">An entity that is not highlighted will not launch a contextual add-in in a card.</span></span>

<span data-ttu-id="a03f0-126">Como uma entidade ou expressão regular que não estiver realçada não iniciará o suplemento contextual, os suplementos devem conter pelo menos um elemento `Rule` com o atributo `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="a03f0-126">Because an entity or regular expression that is not highlighted will not launch a contextual add-in, add-ins must include at least one `Rule` element with the `Highlight` attribute set to `all`.</span></span>

> [!NOTE]
> <span data-ttu-id="a03f0-p107">Os tipos de entidade `EmailAddress` e `Url` não dão suporte ao realce, portanto, não podem ser usados para iniciar um suplemento contextual. No entanto, eles podem ser combinados em um tipo de regra `RuleCollection` como critérios de ativação adicionais.</span><span class="sxs-lookup"><span data-stu-id="a03f0-p107">The `EmailAddress` and `Url` entity types do not support highlighting, so they cannot be used to launch a contextual add-in. They can however be combined in a `RuleCollection` rule type as an additional activation criteria.</span></span>

## <a name="how-to-launch-a-contextual-add-in"></a><span data-ttu-id="a03f0-129">Como iniciar um suplemento contextual</span><span class="sxs-lookup"><span data-stu-id="a03f0-129">How to launch a contextual add-in</span></span>

<span data-ttu-id="a03f0-p108">O usuário inicia o suplemento contextual por meio de texto, tanto uma entidade conhecida quanto uma expressão regular do desenvolvedor. Normalmente, o usuário identifica um suplemento contextual porque a entidade está realçada. O exemplo a seguir mostra como o realce aparece em uma mensagem. Aqui, a entidade (um endereço) está na cor azul e sublinhada com uma linha pontilhada azul. Um usuário inicia o suplemento contextual clicando na entidade realçada.</span><span class="sxs-lookup"><span data-stu-id="a03f0-p108">A user launches a contextual add-in through text, either a known entity or a developer's regular expression. Typically, a user identifies a contextual add-in because the entity is highlighted. The following example shows how highlighting appears in a message. Here the entity (an address) is colored blue and underlined with a dotted blue line. A user launches the contextual add-in by clicking the highlighted entity.</span></span> 

<span data-ttu-id="a03f0-135">**Exemplo de texto com a entidade realçada (um endereço)**</span><span class="sxs-lookup"><span data-stu-id="a03f0-135">**Example of text with highlighted entity (an address)**</span></span>

![Mostra a entidade realçada em um email.](../images/outlook-detected-entity-highlight.png)
    
<span data-ttu-id="a03f0-137">Quando há várias entidades ou suplementos contextuais em uma mensagem, existem algumas regras de interação do usuário:</span><span class="sxs-lookup"><span data-stu-id="a03f0-137">When there are multiple entities or contextual add-ins in a message, there are a few user interaction rules:</span></span>

- <span data-ttu-id="a03f0-138">Se houver várias entidades, o usuário terá que clicar em uma entidade diferente para iniciar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="a03f0-138">If there are multiple entities, the user has to click a different entity to launch the add-in for it.</span></span>
- <span data-ttu-id="a03f0-139">Se uma entidade ativar vários suplementos, cada suplemento abrirá uma nova guia. O usuário alterna entre guias para alternar entre os suplementos. Por exemplo, um nome e um endereço podem acionar um suplemento de telefone e um mapa.</span><span class="sxs-lookup"><span data-stu-id="a03f0-139">If an entity activates multiple add-ins, each add-in opens a new tab. The user switches between tabs to change between add-ins. For example, a name and address might trigger a phone add-in and a map.</span></span>
- <span data-ttu-id="a03f0-p109">Se uma única cadeia de caracteres contiver várias entidades que ativam vários suplementos, toda a cadeia será realçada e um clique na cadeia de caracteres mostra todos os suplementos relevantes à cadeia em guias separadas. Por exemplo, uma cadeia de caracteres que descreve uma reunião proposta em um restaurante pode ativar o suplemento Reunião Sugerida e um suplemento de classificação de restaurantes.</span><span class="sxs-lookup"><span data-stu-id="a03f0-p109">If a single string contains multiple entities that activate multiple add-ins, the entire string is highlighted, and clicking the string shows all add-ins relevant to the string on separate tabs. For example, a string that describes a proposed meeting at a restaurant might activate the Suggested Meeting add-in and a restaurant rating add-in.</span></span>

## <a name="how-a-contextual-add-in-displays"></a><span data-ttu-id="a03f0-142">Como um suplemento contextual é exibido</span><span class="sxs-lookup"><span data-stu-id="a03f0-142">How a contextual add-in displays</span></span>

<span data-ttu-id="a03f0-p110">Um suplemento contextual ativado aparece em um cartão, que é uma janela separada perto a entidade. O cartão normalmente aparecerá abaixo da entidade e centralizado o máximo possível em relação à entidade. Se não houver espaço suficiente embaixo da entidade, o cartão será colocado acima dela. A captura de tela a seguir mostra a entidade realçada e, abaixo dela, um suplemento (Bing Mapas) ativado em um cartão.</span><span class="sxs-lookup"><span data-stu-id="a03f0-p110">An activated contextual add-in appears in a card, which is a separate window near the entity. The card will normally appear below the entity and centered with respect to the entity as much as possible. If there is not enough room below the entity, the card is placed above it. The following screenshot shows the highlighted entity, and below it, an activated add-in (Bing Maps) in a card.</span></span>

<span data-ttu-id="a03f0-147">**Exemplo de um suplemento exibido em um cartão**</span><span class="sxs-lookup"><span data-stu-id="a03f0-147">**Example of an add-in displayed in a card**</span></span>

![Mostra um aplicativo contextual em um cartão.](../images/outlook-detected-entity-card.png)

<span data-ttu-id="a03f0-149">Para fechar o cartão e o suplemento, o usuário deve clicar em algum lugar fora do cartão.</span><span class="sxs-lookup"><span data-stu-id="a03f0-149">To close the card and the add-in, a user clicks anywhere outside of the card.</span></span>

## <a name="current-contextual-add-ins"></a><span data-ttu-id="a03f0-150">Suplementos contextuais atuais</span><span class="sxs-lookup"><span data-stu-id="a03f0-150">Current contextual add-ins</span></span>

<span data-ttu-id="a03f0-151">Os seguintes suplementos contextuais estão instalados por padrão para usuários com os suplementos do Outlook:</span><span class="sxs-lookup"><span data-stu-id="a03f0-151">The following contextual add-ins are installed by default for users with Outlook add-ins:</span></span>

- <span data-ttu-id="a03f0-152">Bing Mapas</span><span class="sxs-lookup"><span data-stu-id="a03f0-152">Bing Maps</span></span> 
- <span data-ttu-id="a03f0-153">Reuniões sugeridas</span><span class="sxs-lookup"><span data-stu-id="a03f0-153">Suggested Meetings</span></span>

## <a name="see-also"></a><span data-ttu-id="a03f0-154">Confira também</span><span class="sxs-lookup"><span data-stu-id="a03f0-154">See also</span></span>

- <span data-ttu-id="a03f0-155">[Suplemento do Outlook: número de ordem da Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (exemplo do suplemento contextual ativado com base em uma correspondência de expressão regular)</span><span class="sxs-lookup"><span data-stu-id="a03f0-155">[Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (sample contextual add-in that activates based on a regular expression match)</span></span>
- [<span data-ttu-id="a03f0-156">Escreva seu primeiro suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="a03f0-156">Write your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="a03f0-157">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="a03f0-157">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="a03f0-158">Objeto Entities</span><span class="sxs-lookup"><span data-stu-id="a03f0-158">Entities object</span></span>](/javascript/api/outlook/office.entities)
