---
title: Elemento IconUrl no arquivo de manifesto
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 44992a3c5f9ceba55b09f4b14e36b5b2935ee669
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771797"
---
# <a name="iconurl-element"></a><span data-ttu-id="a07f2-102">Elemento IconUrl</span><span class="sxs-lookup"><span data-stu-id="a07f2-102">IconUrl element</span></span>

<span data-ttu-id="a07f2-103">Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store.</span><span class="sxs-lookup"><span data-stu-id="a07f2-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="a07f2-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="a07f2-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a07f2-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="a07f2-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="a07f2-106">Pode conter</span><span class="sxs-lookup"><span data-stu-id="a07f2-106">Can contain</span></span>

[<span data-ttu-id="a07f2-107">Override</span><span class="sxs-lookup"><span data-stu-id="a07f2-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="a07f2-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="a07f2-108">Attributes</span></span>

|<span data-ttu-id="a07f2-109">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="a07f2-109">**Attribute**</span></span>|<span data-ttu-id="a07f2-110">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="a07f2-110">**Type**</span></span>|<span data-ttu-id="a07f2-111">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="a07f2-111">**Required**</span></span>|<span data-ttu-id="a07f2-112">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="a07f2-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="a07f2-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="a07f2-113">DefaultValue</span></span>|<span data-ttu-id="a07f2-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a07f2-114">string</span></span>|<span data-ttu-id="a07f2-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="a07f2-115">required</span></span>|<span data-ttu-id="a07f2-116">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="a07f2-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="a07f2-117">Comentários</span><span class="sxs-lookup"><span data-stu-id="a07f2-117">Remarks</span></span>

<span data-ttu-id="a07f2-118">Para \*\*\*\* > um suplemento de email, o ícone é exibido na interface do usuário**gerenciar suplementos** (Outlook) ou **configurações** > **gerenciar suplemento** (Outlook na Web).</span><span class="sxs-lookup"><span data-stu-id="a07f2-118">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="a07f2-119">Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir** > **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="a07f2-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="a07f2-120">Para todos os tipos de suplemento, o ícone também é usado no [AppSource](https://appsource.microsoft.com), se você publicar o suplemento no AppSource.</span><span class="sxs-lookup"><span data-stu-id="a07f2-120">For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.</span></span>

<span data-ttu-id="a07f2-121">A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="a07f2-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="a07f2-122">Para aplicativos de conteúdo e de painel de tarefas, a imagem especificada deve ter 32 x 32 pixels.</span><span class="sxs-lookup"><span data-stu-id="a07f2-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="a07f2-123">Para aplicativos de email, a resolução de imagem recomendada é 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="a07f2-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="a07f2-124">Você também deve especificar um ícone para ser usado com aplicativos host do Office executados em telas de DPI alto que utilizam o elemento [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="a07f2-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="a07f2-125">Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="a07f2-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="a07f2-126">Não há suporte atualmente para `IconUrl` a alteração do valor do elemento no tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="a07f2-126">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>