---
title: Conjuntos de requisitos da API JavaScript do PowerPoint
description: Saiba mais sobre os conjuntos de requisitos da API JavaScript do PowerPoint.
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: cf9ab510e4b35a140c77ee958279cb85a2189fa2
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774723"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="0aaa1-103">Conjuntos de requisitos da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0aaa1-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="0aaa1-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="0aaa1-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="0aaa1-107">A tabela a seguir lista os conjuntos de requisitos do PowerPoint, os aplicativos do cliente Office que oferecem suporte a esses conjuntos de requisitos e as versões de compilação ou datas de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="0aaa1-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="0aaa1-108">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="0aaa1-108">Requirement set</span></span>  |  <span data-ttu-id="0aaa1-109">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="0aaa1-109">Office on Windows</span></span><br><span data-ttu-id="0aaa1-110">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="0aaa1-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="0aaa1-111">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="0aaa1-111">Office on iPad</span></span><br><span data-ttu-id="0aaa1-112">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="0aaa1-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="0aaa1-113">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="0aaa1-113">Office on Mac</span></span><br><span data-ttu-id="0aaa1-114">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="0aaa1-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="0aaa1-115">Office na Web</span><span class="sxs-lookup"><span data-stu-id="0aaa1-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| [<span data-ttu-id="0aaa1-116">Visualização</span><span class="sxs-lookup"><span data-stu-id="0aaa1-116">Preview</span></span>](powerpoint-preview-apis.md)  | <span data-ttu-id="0aaa1-117">Use a versão mais recente do Office para experimentar APIs de visualização (pode ser necessário ingressar no [Programa Office Insider](https://insider.office.com)).</span><span class="sxs-lookup"><span data-stu-id="0aaa1-117">Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com)).</span></span> |
| <span data-ttu-id="0aaa1-118">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="0aaa1-118">PowerPointApi 1.1</span></span> | <span data-ttu-id="0aaa1-119">Versão 1810 (Build 11001.20074) ou posterior</span><span class="sxs-lookup"><span data-stu-id="0aaa1-119">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="0aaa1-120">2.17 ou posterior</span><span class="sxs-lookup"><span data-stu-id="0aaa1-120">2.17 or later</span></span> | <span data-ttu-id="0aaa1-121">16.19 ou posterior</span><span class="sxs-lookup"><span data-stu-id="0aaa1-121">16.19 or later</span></span> | <span data-ttu-id="0aaa1-122">Outubro de 2018</span><span class="sxs-lookup"><span data-stu-id="0aaa1-122">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="0aaa1-123">Versões do Office e números de build</span><span class="sxs-lookup"><span data-stu-id="0aaa1-123">Office versions and build numbers</span></span>

<span data-ttu-id="0aaa1-124">Para saber mais sobre as versões do Office e os números de build, confira:</span><span class="sxs-lookup"><span data-stu-id="0aaa1-124">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="0aaa1-125">API JavaScript do PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="0aaa1-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="0aaa1-126">O PowerPoint JavaScript API 1.1 contém uma [única API para criar uma nova apresentação](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span><span class="sxs-lookup"><span data-stu-id="0aaa1-126">PowerPoint JavaScript API 1.1 contains a [single API to create a new presentation](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span></span> <span data-ttu-id="0aaa1-127">Para obter detalhes sobre a API, confira [Criar uma apresentação](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span><span class="sxs-lookup"><span data-stu-id="0aaa1-127">For details about the API, see [Create a presentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span></span>

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a><span data-ttu-id="0aaa1-128">Como usar os conjuntos de requisitos do PowerPoint em tempo de execução e no manifesto</span><span class="sxs-lookup"><span data-stu-id="0aaa1-128">How to use PowerPoint requirement sets at runtime and in the manifest</span></span>

> [!NOTE]
> <span data-ttu-id="0aaa1-129">Esta seção pressupõe que você esteja familiarizado com a visão geral dos conjuntos de requisitos em [Versões e conjuntos de requisitos do Office](../../develop/office-versions-and-requirement-sets.md) e [Especificar aplicativos do Office e requisitos de API](../../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="0aaa1-129">This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md) and [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).</span></span>

<span data-ttu-id="0aaa1-130">Os conjuntos de requisitos são grupos nomeados de membros da API.</span><span class="sxs-lookup"><span data-stu-id="0aaa1-130">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="0aaa1-131">Um Suplemento do Office pode executar uma verificação de tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um aplicativo do Office dá suporte às APIs necessárias ao suplemento.</span><span class="sxs-lookup"><span data-stu-id="0aaa1-131">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="0aaa1-132">Verificando o suporte ao conjunto de requisitos no tempo de execução</span><span class="sxs-lookup"><span data-stu-id="0aaa1-132">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="0aaa1-133">O exemplo de código a seguir mostra como determinar se o aplicativo do Office, onde o suplemento está em execução, dá suporte ao conjunto de requisitos da API especificado.</span><span class="sxs-lookup"><span data-stu-id="0aaa1-133">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="0aaa1-134">Definindo o suporte ao conjunto de requisitos no manifesto</span><span class="sxs-lookup"><span data-stu-id="0aaa1-134">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="0aaa1-135">Você pode usar o [elemento Requirements](../manifest/requirements.md) no manifesto do suplemento para especificar os conjuntos de requisitos mínimos e/ou os métodos de API exigidos pelo suplemento para ser ativado.</span><span class="sxs-lookup"><span data-stu-id="0aaa1-135">You can use the [Requirements element](../manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="0aaa1-136">Se a plataforma ou o aplicativo do Office não for compatível com os conjuntos de requisitos ou métodos de API especificados no `Requirements` elemento do manifesto, o suplemento não será executado nesse aplicativo ou plataforma, e não será exibido na lista de suplementos mostrados no **Meus suplementos** . Se o seu suplemento exige um conjunto específico de requisitos para funcionalidade total, mas pode fornecer um valor mesmo para os usuários nas plataformas que não têm suporte para o conjunto de requisitos, recomendamos verificar o suporte a requisitos no tempo de execução conforme descrito acima, em vez de definir o suporte ao conjunto de requisitos no manifesto.</span><span class="sxs-lookup"><span data-stu-id="0aaa1-136">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins** . If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that do not support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.</span></span>

<span data-ttu-id="0aaa1-137">O exemplo de código a seguir mostra o elemento `Requirements` em um manifesto de suplemento que especifica que o suplemento deve ser carregado em todos os aplicativos do cliente do Office que oferecem suporte ao conjunto de requisitos da versão 1.1 ou superior do PowerPointApi.</span><span class="sxs-lookup"><span data-stu-id="0aaa1-137">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support PowerPointApi requirement set version 1.1 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="0aaa1-138">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="0aaa1-138">Office Common API requirement sets</span></span>

<span data-ttu-id="0aaa1-139">A maior parte da funcionalidade do suplemento do PowerPoint vem do conjunto de APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="0aaa1-139">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="0aaa1-140">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="0aaa1-140">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0aaa1-141">Confira também</span><span class="sxs-lookup"><span data-stu-id="0aaa1-141">See also</span></span>

- [<span data-ttu-id="0aaa1-142">Documentação de Referência da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0aaa1-142">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="0aaa1-143">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="0aaa1-143">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="0aaa1-144">Especificar requisitos da API e de aplicativos do Office</span><span class="sxs-lookup"><span data-stu-id="0aaa1-144">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="0aaa1-145">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="0aaa1-145">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
