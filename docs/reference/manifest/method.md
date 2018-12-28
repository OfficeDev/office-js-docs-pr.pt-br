---
title: Elemento Method no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: fded84344182bb45597b00a794f18defaa44d3b3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432820"
---
# <a name="method-element"></a><span data-ttu-id="94655-102">Elemento Method</span><span class="sxs-lookup"><span data-stu-id="94655-102">Method element</span></span>

<span data-ttu-id="94655-103">Especifica um método individual a partir da API do JavaScript para Office que o Suplemento do Office exige para ativar.</span><span class="sxs-lookup"><span data-stu-id="94655-103">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="94655-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="94655-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="94655-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="94655-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="94655-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="94655-106">Contained in</span></span>

[<span data-ttu-id="94655-107">Methods</span><span class="sxs-lookup"><span data-stu-id="94655-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="94655-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="94655-108">Attributes</span></span>

|<span data-ttu-id="94655-109">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="94655-109">**Attribute**</span></span>|<span data-ttu-id="94655-110">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="94655-110">**Type**</span></span>|<span data-ttu-id="94655-111">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="94655-111">**Required**</span></span>|<span data-ttu-id="94655-112">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="94655-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="94655-113">Nome</span><span class="sxs-lookup"><span data-stu-id="94655-113">Name</span></span>|<span data-ttu-id="94655-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="94655-114">string</span></span>|<span data-ttu-id="94655-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="94655-115">required</span></span>|<span data-ttu-id="94655-p101">Especifica o nome do método necessário qualificado com seu objeto pai. Por exemplo, para especificar o método **getSelectedDataAsync**, você deve especificar `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="94655-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="94655-118">Comentários</span><span class="sxs-lookup"><span data-stu-id="94655-118">Remarks</span></span>

<span data-ttu-id="94655-119">Os elementos **Method** e **Methods** não têm suporte nos suplementos de email. Para saber mais sobre conjuntos de requisitos, consulte [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="94655-119">The  Methods and Method elements aren't supported by mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="94655-120">Como não há forma de especificar o requisito de versão mínimo de métodos individuais, para verificar se um método está disponível no tempo de execução, você também deve usar uma instrução **if** ao chamar esse método no script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="94655-120">Important  Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an  **if** statement when calling that method in the script of your add-in. For more information about how to do this, see Understanding the JavaScript API for Office.</span></span> <span data-ttu-id="94655-121">Para saber mais sobre como fazer isso, consulte [Noções básicas da API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="94655-121">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

