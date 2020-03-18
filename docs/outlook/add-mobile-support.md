---
title: Adicionar suporte móvel a um suplemento do Outlook
description: A adição de suporte para o Outlook Mobile requer atualização do manifesto do suplemento e, possivelmente, a alteração do código para cenários móveis.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 31f58102129ae207da55839f7b48cc8a060645ad
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720852"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a><span data-ttu-id="504de-103">Adicionar suporte para comandos de suplementos para Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="504de-103">Add support for add-in commands for Outlook Mobile</span></span>

<span data-ttu-id="504de-104">O uso de comandos de suplemento no Outlook Mobile permite que os usuários acessem a mesma funcionalidade (com algumas [limitações](#code-considerations)) já existentes no Outlook na Web, no Windows e no Mac.</span><span class="sxs-lookup"><span data-stu-id="504de-104">Using add-in commands in Outlook Mobile allows your users to access the same functionality (with some [limitations](#code-considerations)) that they already have in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="504de-105">A adição de suporte para o Outlook Mobile requer atualização do manifesto do suplemento e, possivelmente, a alteração do código para cenários móveis.</span><span class="sxs-lookup"><span data-stu-id="504de-105">Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.</span></span>

## <a name="updating-the-manifest"></a><span data-ttu-id="504de-106">Atualização do manifesto</span><span class="sxs-lookup"><span data-stu-id="504de-106">Updating the manifest</span></span>

<span data-ttu-id="504de-p102">A primeira etapa para habilitar os comandos de suplemento no Outlook Mobile é defini-los no manifesto do suplemento. O esquema [VersionOverrides](../reference/manifest/versionoverrides.md) versão 1.1 define um novo fator forma para dispositivos móveis, o [MobileFormFactor](../reference/manifest/mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="504de-p102">The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](../reference/manifest/mobileformfactor.md).</span></span>

<span data-ttu-id="504de-p103">Esse elemento contém todas as informações para carregar o suplemento em clientes móveis. Isso permite que você defina elementos de interface completamente diferentes e arquivos JavaScript para a experiência móvel.</span><span class="sxs-lookup"><span data-stu-id="504de-p103">This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.</span></span>

<span data-ttu-id="504de-111">O exemplo a seguir mostra um único botão de painel de `MobileFormFactor` tarefas em um elemento.</span><span class="sxs-lookup"><span data-stu-id="504de-111">The following example shows a single task pane button in a `MobileFormFactor` element.</span></span>

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

<span data-ttu-id="504de-112">Isso é muito semelhante aos elementos que aparecem em um elemento [DesktopFormFactor](../reference/manifest/desktopformfactor.md), com algumas diferenças importantes.</span><span class="sxs-lookup"><span data-stu-id="504de-112">This is very similar to the elements that appear in a [DesktopFormFactor](../reference/manifest/desktopformfactor.md) element, with some notable differences.</span></span>

- <span data-ttu-id="504de-113">O elemento [OfficeTab](../reference/manifest/officetab.md) não é usado.</span><span class="sxs-lookup"><span data-stu-id="504de-113">The [OfficeTab](../reference/manifest/officetab.md) element is not used.</span></span>
- <span data-ttu-id="504de-p104">O elemento [ExtensionPoint](../reference/manifest/extensionpoint.md) deve ter apenas um elemento filho. Se o suplemento apenas adiciona um botão, o elemento filho deve ser um elemento [Control](../reference/manifest/control.md). Se o suplemento adiciona mais de um botão, o elemento filho deve ser um elemento [Group](../reference/manifest/group.md) que contém vários elementos `Control`.</span><span class="sxs-lookup"><span data-stu-id="504de-p104">The [ExtensionPoint](../reference/manifest/extensionpoint.md) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](../reference/manifest/control.md) element. If the add-in adds more than one button, the child element should be a [Group](../reference/manifest/group.md) element that contains multiple `Control` elements.</span></span>
- <span data-ttu-id="504de-117">Não há nenhum tipo `Menu` equivalente ao elemento `Control`.</span><span class="sxs-lookup"><span data-stu-id="504de-117">There is no `Menu` type equivalent for the `Control` element.</span></span>
- <span data-ttu-id="504de-118">O elemento [Supertip](../reference/manifest/supertip.md) não é usado.</span><span class="sxs-lookup"><span data-stu-id="504de-118">The [Supertip](../reference/manifest/supertip.md) element is not used.</span></span>
- <span data-ttu-id="504de-p105">Os tamanhos de ícone obrigatórios são diferentes. Suplementos móveis devem, no mínimo, dar suporte a ícones de 25 x 25, 32 x 32 e 48 x 48 pixels.</span><span class="sxs-lookup"><span data-stu-id="504de-p105">The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.</span></span>

## <a name="code-considerations"></a><span data-ttu-id="504de-121">Considerações sobre código</span><span class="sxs-lookup"><span data-stu-id="504de-121">Code considerations</span></span>

<span data-ttu-id="504de-122">Criar um suplemento para o Mobile traz algumas considerações adicionais.</span><span class="sxs-lookup"><span data-stu-id="504de-122">Designing an add-in for mobile introduces some additional considerations.</span></span>

### <a name="use-rest-instead-of-exchange-web-services"></a><span data-ttu-id="504de-123">Usar REST em vez de Serviços Web do Exchange</span><span class="sxs-lookup"><span data-stu-id="504de-123">Use REST instead of Exchange Web Services</span></span>

<span data-ttu-id="504de-p106">O método [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) não é suportado no Outlook Mobile. Os suplementos devem preferir obter as informações da API Office.js sempre que possível. Se os suplementos exigem informações que não são expostas pela API Office.js devem usar as [APIs REST do Outlook](/outlook/rest/) para acessar as caixas de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="504de-p106">The [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.</span></span>

<span data-ttu-id="504de-127">O conjunto de requisitos de caixa de correio 1,5 introduziu uma nova versão do [Office. Context. Mailbox. getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) que pode solicitar um token de acesso compatível com as APIs REST e uma nova propriedade [Office. Context. Mailbox. restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) que pode ser usada para localizar o ponto de extremidade da API REST para o usuário.</span><span class="sxs-lookup"><span data-stu-id="504de-127">Mailbox requirement set 1.5 introduced a new version of [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) that can request an access token compatible with the REST APIs, and a new [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property that can be used to find the REST API endpoint for the user.</span></span>

### <a name="pinch-zoom"></a><span data-ttu-id="504de-128">Pinçar e zoom</span><span class="sxs-lookup"><span data-stu-id="504de-128">Pinch zoom</span></span>

<span data-ttu-id="504de-p107">Por padrão, os usuários podem usar o gesto de “pinçar/zoom” para aplicar zoom aos painéis de tarefas. Se isso não fizer sentido em seu cenário, desative esse recurso em seu HTML.</span><span class="sxs-lookup"><span data-stu-id="504de-p107">By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.</span></span>

### <a name="close-task-panes"></a><span data-ttu-id="504de-131">Fechar painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="504de-131">Close task panes</span></span>

<span data-ttu-id="504de-p108">Nos Outlook Mobile, os painéis de tarefa ocupam a tela inteira e, por padrão, exigem que o usuário os feche para retornar à mensagem. Considere o uso do método [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) para fechar o painel de tarefas quando seu cenário estiver concluído.</span><span class="sxs-lookup"><span data-stu-id="504de-p108">In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) method to close the task pane when your scenario is complete.</span></span>

### <a name="compose-mode-and-appointments"></a><span data-ttu-id="504de-134">Modo de redação e compromissos</span><span class="sxs-lookup"><span data-stu-id="504de-134">Compose mode and appointments</span></span>

<span data-ttu-id="504de-p109">Atualmente, os suplementos do Outlook Mobile dão suporte à ativação apenas durante a leitura de mensagens. Os suplementos não são ativados ao redigir mensagens ou ao exibir ou redigir compromissos.</span><span class="sxs-lookup"><span data-stu-id="504de-p109">Currently add-ins in Outlook Mobile only support activation when reading messages. Add-ins are not activated when composing messages or when viewing or composing appointments.</span></span>

### <a name="unsupported-apis"></a><span data-ttu-id="504de-137">APIs sem suporte</span><span class="sxs-lookup"><span data-stu-id="504de-137">Unsupported APIs</span></span>

<span data-ttu-id="504de-138">As APIs introduzidas no conjunto de requisitos 1,6 ou posterior não são suportadas pelo Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="504de-138">APIs introduced in requirement set 1.6 or later are not supported by Outlook Mobile.</span></span> <span data-ttu-id="504de-139">As seguintes APIs de conjuntos de requisitos anteriores também não são suportadas.</span><span class="sxs-lookup"><span data-stu-id="504de-139">The following APIs from earlier requirement sets are also not supported.</span></span>

  - [<span data-ttu-id="504de-140">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="504de-140">Office.context.officeTheme</span></span>](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
  - [<span data-ttu-id="504de-141">Office.context.mailbox.ewsUrl</span><span class="sxs-lookup"><span data-stu-id="504de-141">Office.context.mailbox.ewsUrl</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
  - [<span data-ttu-id="504de-142">Office.context.mailbox.convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="504de-142">Office.context.mailbox.convertToEwsId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="504de-143">Office.context.mailbox.convertToRestId</span><span class="sxs-lookup"><span data-stu-id="504de-143">Office.context.mailbox.convertToRestId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="504de-144">Office.context.mailbox.displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="504de-144">Office.context.mailbox.displayAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="504de-145">Office.context.mailbox.displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="504de-145">Office.context.mailbox.displayMessageForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="504de-146">Office.context.mailbox.displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="504de-146">Office.context.mailbox.displayNewAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="504de-147">Office.context.mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="504de-147">Office.context.mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="504de-148">Office.context.mailbox.item.dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="504de-148">Office.context.mailbox.item.dateTimeModified</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
  - [<span data-ttu-id="504de-149">Office.context.mailbox.item.displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="504de-149">Office.context.mailbox.item.displayReplyAllForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="504de-150">Office.context.mailbox.item.displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="504de-150">Office.context.mailbox.item.displayReplyForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="504de-151">Office.context.mailbox.item.getEntities</span><span class="sxs-lookup"><span data-stu-id="504de-151">Office.context.mailbox.item.getEntities</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="504de-152">Office.context.mailbox.item.getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="504de-152">Office.context.mailbox.item.getEntitiesByType</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="504de-153">Office.context.mailbox.item.getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="504de-153">Office.context.mailbox.item.getFilteredEntitiesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="504de-154">Office.context.mailbox.item.getRegexMatches</span><span class="sxs-lookup"><span data-stu-id="504de-154">Office.context.mailbox.item.getRegexMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="504de-155">Office.context.mailbox.item.getRegexMatchesByName</span><span class="sxs-lookup"><span data-stu-id="504de-155">Office.context.mailbox.item.getRegexMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a><span data-ttu-id="504de-156">Também confira</span><span class="sxs-lookup"><span data-stu-id="504de-156">See also</span></span>

[<span data-ttu-id="504de-157">Suporte ao conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="504de-157">Requirement set support</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)