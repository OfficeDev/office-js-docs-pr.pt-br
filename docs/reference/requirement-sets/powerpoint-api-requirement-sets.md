---
title: Conjuntos de requisitos da API JavaScript do PowerPoint
description: Saiba mais sobre os conjuntos de requisitos da API JavaScript do PowerPoint.
ms.date: 01/08/2021
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 63f11f1810b38471a27766843f512da193394838
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49840080"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="5cf5d-103">Conjuntos de requisitos da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5cf5d-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="5cf5d-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="5cf5d-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="5cf5d-107">A tabela a seguir lista os conjuntos de requisitos do PowerPoint, os aplicativos do cliente Office que oferecem suporte a esses conjuntos de requisitos e as versões de compilação ou datas de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="5cf5d-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="5cf5d-108">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="5cf5d-108">Requirement set</span></span>  |  <span data-ttu-id="5cf5d-109">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="5cf5d-109">Office on Windows</span></span><br><span data-ttu-id="5cf5d-110">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="5cf5d-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="5cf5d-111">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="5cf5d-111">Office on iPad</span></span><br><span data-ttu-id="5cf5d-112">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="5cf5d-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="5cf5d-113">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="5cf5d-113">Office on Mac</span></span><br><span data-ttu-id="5cf5d-114">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="5cf5d-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="5cf5d-115">Office na Web</span><span class="sxs-lookup"><span data-stu-id="5cf5d-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| [<span data-ttu-id="5cf5d-116">PowerPointApi 1.2</span><span class="sxs-lookup"><span data-stu-id="5cf5d-116">PowerPointApi 1.2</span></span>](powerpoint-api-1-2-requirement-set.md)  | <span data-ttu-id="5cf5d-117">Versão 2011 (Compilação 13426.20184) ou superior</span><span class="sxs-lookup"><span data-stu-id="5cf5d-117">Version 2011 (Build 13426.20184) or later</span></span>| <span data-ttu-id="5cf5d-118">ainda não</span><span class="sxs-lookup"><span data-stu-id="5cf5d-118">not yet</span></span><br><span data-ttu-id="5cf5d-119">com suporte</span><span class="sxs-lookup"><span data-stu-id="5cf5d-119">supported</span></span> | <span data-ttu-id="5cf5d-120">16.43 ou superior</span><span class="sxs-lookup"><span data-stu-id="5cf5d-120">16.43 or later</span></span> | <span data-ttu-id="5cf5d-121">outubro de 2020</span><span class="sxs-lookup"><span data-stu-id="5cf5d-121">October 2020</span></span> |
| [<span data-ttu-id="5cf5d-122">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="5cf5d-122">PowerPointApi 1.1</span></span>](powerpoint-api-1-1-requirement-set.md) | <span data-ttu-id="5cf5d-123">Versão 1810 (Build 11001.20074) ou posterior</span><span class="sxs-lookup"><span data-stu-id="5cf5d-123">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="5cf5d-124">2.17 ou posterior</span><span class="sxs-lookup"><span data-stu-id="5cf5d-124">2.17 or later</span></span> | <span data-ttu-id="5cf5d-125">16.19 ou posterior</span><span class="sxs-lookup"><span data-stu-id="5cf5d-125">16.19 or later</span></span> | <span data-ttu-id="5cf5d-126">Outubro de 2018</span><span class="sxs-lookup"><span data-stu-id="5cf5d-126">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="5cf5d-127">Versões do Office e números de build</span><span class="sxs-lookup"><span data-stu-id="5cf5d-127">Office versions and build numbers</span></span>

<span data-ttu-id="5cf5d-128">Para saber mais sobre as versões do Office e os números de build, confira:</span><span class="sxs-lookup"><span data-stu-id="5cf5d-128">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="5cf5d-129">API JavaScript do PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="5cf5d-129">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="5cf5d-130">O PowerPoint JavaScript API 1.1 contém uma [única API para criar uma nova apresentação](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span><span class="sxs-lookup"><span data-stu-id="5cf5d-130">PowerPoint JavaScript API 1.1 contains a [single API to create a new presentation](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span></span> <span data-ttu-id="5cf5d-131">Para obter detalhes sobre a API, confira [Criar uma apresentação](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span><span class="sxs-lookup"><span data-stu-id="5cf5d-131">For details about the API, see [Create a presentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span></span>

## <a name="powerpoint-javascript-api-12"></a><span data-ttu-id="5cf5d-132">API JavaScript do PowerPoint 1.2</span><span class="sxs-lookup"><span data-stu-id="5cf5d-132">PowerPoint JavaScript API 1.2</span></span>

<span data-ttu-id="5cf5d-133">A API JavaScript do PowerPoint 1.2 adiciona suporte para inserir slides de outra apresentação do PowerPoint na apresentação atual e para excluir slides.</span><span class="sxs-lookup"><span data-stu-id="5cf5d-133">PowerPoint JavaScript API 1.2 adds support for inserting slides from another PowerPoint presentation into the current presentation and for deleting slides.</span></span> <span data-ttu-id="5cf5d-134">Para obter detalhes sobre as APIs, consulte [Inserir e excluir slides em uma apresentação do PowerPoint](../../powerpoint/insert-slides-into-presentation.md).</span><span class="sxs-lookup"><span data-stu-id="5cf5d-134">For details about the APIs, see [Insert and delete slides in a PowerPoint presentation](../../powerpoint/insert-slides-into-presentation.md).</span></span>

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a><span data-ttu-id="5cf5d-135">Como usar os conjuntos de requisitos do PowerPoint em tempo de execução e no manifesto</span><span class="sxs-lookup"><span data-stu-id="5cf5d-135">How to use PowerPoint requirement sets at runtime and in the manifest</span></span>

> [!NOTE]
> <span data-ttu-id="5cf5d-136">Esta seção pressupõe que você esteja familiarizado com a visão geral dos conjuntos de requisitos em [Versões e conjuntos de requisitos do Office](../../develop/office-versions-and-requirement-sets.md) e [Especificar aplicativos do Office e requisitos de API](../../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="5cf5d-136">This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md) and [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).</span></span>

<span data-ttu-id="5cf5d-137">Os conjuntos de requisitos são grupos nomeados de membros da API.</span><span class="sxs-lookup"><span data-stu-id="5cf5d-137">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="5cf5d-138">Um Suplemento do Office pode executar uma verificação de tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um aplicativo do Office dá suporte às APIs necessárias ao suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cf5d-138">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="5cf5d-139">Verificando o suporte ao conjunto de requisitos no tempo de execução</span><span class="sxs-lookup"><span data-stu-id="5cf5d-139">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="5cf5d-140">O exemplo de código a seguir mostra como determinar se o aplicativo do Office, onde o suplemento está em execução, dá suporte ao conjunto de requisitos da API especificado.</span><span class="sxs-lookup"><span data-stu-id="5cf5d-140">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="5cf5d-141">Definindo o suporte ao conjunto de requisitos no manifesto</span><span class="sxs-lookup"><span data-stu-id="5cf5d-141">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="5cf5d-142">Você pode usar o [elemento Requirements](../manifest/requirements.md) no manifesto do suplemento para especificar os conjuntos de requisitos mínimos e/ou os métodos de API exigidos pelo suplemento para ser ativado.</span><span class="sxs-lookup"><span data-stu-id="5cf5d-142">You can use the [Requirements element](../manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="5cf5d-143">Se o aplicativo ou plataforma do Office não suportar os conjuntos de requisitos ou métodos de API que são especificados no elemento `Requirements` do manifesto, o suplemento não será executado nesse aplicativo ou plataforma e não será exibido no lista de suplementos que são mostrados em **Meus suplementos**. Se o seu suplemento requer um conjunto de requisitos específico para funcionalidade total, mas pode fornecer valor até mesmo para usuários em plataformas que não oferecem suporte ao conjunto de requisitos, recomendamos que você verifique o suporte ao requisito no tempo de execução conforme descrito acima, em vez de definir o suporte ao conjunto de requisitos no manifesto.</span><span class="sxs-lookup"><span data-stu-id="5cf5d-143">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**. If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that don't support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.</span></span>

<span data-ttu-id="5cf5d-144">O exemplo de código a seguir mostra o elemento `Requirements` em um manifesto de suplemento que especifica que o suplemento deve ser carregado em todos os aplicativos do cliente do Office que oferecem suporte ao conjunto de requisitos da versão 1.1 ou superior do PowerPointApi.</span><span class="sxs-lookup"><span data-stu-id="5cf5d-144">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support PowerPointApi requirement set version 1.1 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="5cf5d-145">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="5cf5d-145">Office Common API requirement sets</span></span>

<span data-ttu-id="5cf5d-146">A maior parte da funcionalidade do suplemento do PowerPoint vem do conjunto de APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="5cf5d-146">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="5cf5d-147">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="5cf5d-147">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="5cf5d-148">Confira também</span><span class="sxs-lookup"><span data-stu-id="5cf5d-148">See also</span></span>

- [<span data-ttu-id="5cf5d-149">Documentação de Referência da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5cf5d-149">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="5cf5d-150">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="5cf5d-150">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="5cf5d-151">Especificar requisitos da API e de aplicativos do Office</span><span class="sxs-lookup"><span data-stu-id="5cf5d-151">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="5cf5d-152">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5cf5d-152">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
