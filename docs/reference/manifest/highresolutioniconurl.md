---
title: Elemento HighResolutionIconUrl no arquivo de manifesto
description: Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store em telas de DPI alto.
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 77675e768895a568bdfee97fc4d5006e1e890937
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641351"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="d2b95-103">Elemento HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="d2b95-103">HighResolutionIconUrl element</span></span>

<span data-ttu-id="d2b95-104">Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store em telas de DPI alto.</span><span class="sxs-lookup"><span data-stu-id="d2b95-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="d2b95-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="d2b95-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d2b95-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="d2b95-106">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="d2b95-107">Pode conter</span><span class="sxs-lookup"><span data-stu-id="d2b95-107">Can contain</span></span>

[<span data-ttu-id="d2b95-108">Override</span><span class="sxs-lookup"><span data-stu-id="d2b95-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="d2b95-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="d2b95-109">Attributes</span></span>

|<span data-ttu-id="d2b95-110">Atributo</span><span class="sxs-lookup"><span data-stu-id="d2b95-110">Attribute</span></span>|<span data-ttu-id="d2b95-111">Tipo</span><span class="sxs-lookup"><span data-stu-id="d2b95-111">Type</span></span>|<span data-ttu-id="d2b95-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="d2b95-112">Required</span></span>|<span data-ttu-id="d2b95-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="d2b95-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d2b95-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="d2b95-114">DefaultValue</span></span>|<span data-ttu-id="d2b95-115">cadeia de caracteres (URL)</span><span class="sxs-lookup"><span data-stu-id="d2b95-115">string (URL)</span></span>|<span data-ttu-id="d2b95-116">obrigatório</span><span class="sxs-lookup"><span data-stu-id="d2b95-116">required</span></span>|<span data-ttu-id="d2b95-117">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="d2b95-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="d2b95-118">Comentários</span><span class="sxs-lookup"><span data-stu-id="d2b95-118">Remarks</span></span>

<span data-ttu-id="d2b95-119">Para um suplemento de email, o ícone é exibido na **File**  >  interface do usuário**gerenciar suplementos** .</span><span class="sxs-lookup"><span data-stu-id="d2b95-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="d2b95-120">Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir** > **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="d2b95-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="d2b95-121">A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="d2b95-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="d2b95-122">Para aplicativos do painel de tarefas e de conteúdo, a resolução de imagem recomendada é 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="d2b95-122">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="d2b95-123">Para aplicativos de email, a imagem deve ter 128 x 128 pixels.</span><span class="sxs-lookup"><span data-stu-id="d2b95-123">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="d2b95-124">Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="d2b95-124">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
