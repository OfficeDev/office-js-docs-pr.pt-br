---
title: Elemento HighResolutionIconUrl no arquivo de manifesto
description: Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store em telas de DPI alto.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 42a7ebf0e02eb365962b574821d5a7004a8b867f
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604637"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="26831-103">Elemento HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="26831-103">HighResolutionIconUrl element</span></span>

<span data-ttu-id="26831-104">Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store em telas de DPI alto.</span><span class="sxs-lookup"><span data-stu-id="26831-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="26831-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="26831-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="26831-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="26831-106">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="26831-107">Pode conter</span><span class="sxs-lookup"><span data-stu-id="26831-107">Can contain</span></span>

[<span data-ttu-id="26831-108">Override</span><span class="sxs-lookup"><span data-stu-id="26831-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="26831-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="26831-109">Attributes</span></span>

|<span data-ttu-id="26831-110">Atributo</span><span class="sxs-lookup"><span data-stu-id="26831-110">Attribute</span></span>|<span data-ttu-id="26831-111">Tipo</span><span class="sxs-lookup"><span data-stu-id="26831-111">Type</span></span>|<span data-ttu-id="26831-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="26831-112">Required</span></span>|<span data-ttu-id="26831-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="26831-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="26831-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="26831-114">DefaultValue</span></span>|<span data-ttu-id="26831-115">cadeia de caracteres (URL)</span><span class="sxs-lookup"><span data-stu-id="26831-115">string (URL)</span></span>|<span data-ttu-id="26831-116">obrigatório</span><span class="sxs-lookup"><span data-stu-id="26831-116">required</span></span>|<span data-ttu-id="26831-117">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="26831-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="26831-118">Comentários</span><span class="sxs-lookup"><span data-stu-id="26831-118">Remarks</span></span>

<span data-ttu-id="26831-119">Para um complemento de email, o ícone é exibido na interface do usuário Gerenciar arquivos de  >  **complementos.**</span><span class="sxs-lookup"><span data-stu-id="26831-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="26831-120">Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir** > **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="26831-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="26831-121">A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="26831-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="26831-122">Para aplicativos de conteúdo e painel de tarefas, a resolução de imagem deve ser de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="26831-122">For content and task pane apps, the image resolution must be 64 x 64 pixels.</span></span> <span data-ttu-id="26831-123">Para aplicativos de email, a imagem deve ter 128 x 128 pixels.</span><span class="sxs-lookup"><span data-stu-id="26831-123">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="26831-124">Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="26831-124">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
