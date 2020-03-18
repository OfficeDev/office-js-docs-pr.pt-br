---
title: Elemento IconUrl no arquivo de manifesto
description: O elemento IconUrl especifica a URL da imagem que representa o suplemento do Office no UX de inserção e na Office Store.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: a345971e32e64557005c8d01519589f4be5fb7d7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718080"
---
# <a name="iconurl-element"></a><span data-ttu-id="4a5f1-103">Elemento IconUrl</span><span class="sxs-lookup"><span data-stu-id="4a5f1-103">IconUrl element</span></span>

<span data-ttu-id="4a5f1-104">Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store.</span><span class="sxs-lookup"><span data-stu-id="4a5f1-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="4a5f1-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="4a5f1-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4a5f1-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="4a5f1-106">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="4a5f1-107">Pode conter</span><span class="sxs-lookup"><span data-stu-id="4a5f1-107">Can contain</span></span>

[<span data-ttu-id="4a5f1-108">Override</span><span class="sxs-lookup"><span data-stu-id="4a5f1-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="4a5f1-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="4a5f1-109">Attributes</span></span>

|<span data-ttu-id="4a5f1-110">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="4a5f1-110">**Attribute**</span></span>|<span data-ttu-id="4a5f1-111">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="4a5f1-111">**Type**</span></span>|<span data-ttu-id="4a5f1-112">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="4a5f1-112">**Required**</span></span>|<span data-ttu-id="4a5f1-113">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="4a5f1-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="4a5f1-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="4a5f1-114">DefaultValue</span></span>|<span data-ttu-id="4a5f1-115">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4a5f1-115">string</span></span>|<span data-ttu-id="4a5f1-116">obrigatório</span><span class="sxs-lookup"><span data-stu-id="4a5f1-116">required</span></span>|<span data-ttu-id="4a5f1-117">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="4a5f1-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="4a5f1-118">Comentários</span><span class="sxs-lookup"><span data-stu-id="4a5f1-118">Remarks</span></span>

<span data-ttu-id="4a5f1-119">Para > um suplemento de email, o ícone é exibido na **interface do usuário\*\*\*\*gerenciar suplementos** (Outlook) ou **configurações** > **gerenciar suplemento** (Outlook na Web).</span><span class="sxs-lookup"><span data-stu-id="4a5f1-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="4a5f1-120">Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir** > **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="4a5f1-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="4a5f1-121">Para todos os tipos de suplemento, o ícone também é usado no [AppSource](https://appsource.microsoft.com), se você publicar o suplemento no AppSource.</span><span class="sxs-lookup"><span data-stu-id="4a5f1-121">For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.</span></span>

<span data-ttu-id="4a5f1-122">A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="4a5f1-122">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="4a5f1-123">Para aplicativos de conteúdo e de painel de tarefas, a imagem especificada deve ter 32 x 32 pixels.</span><span class="sxs-lookup"><span data-stu-id="4a5f1-123">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="4a5f1-124">Para aplicativos de email, a resolução de imagem recomendada é 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="4a5f1-124">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="4a5f1-125">Você também deve especificar um ícone para ser usado com aplicativos host do Office executados em telas de DPI alto que utilizam o elemento [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="4a5f1-125">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="4a5f1-126">Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="4a5f1-126">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="4a5f1-127">Não há suporte atualmente para `IconUrl` a alteração do valor do elemento no tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="4a5f1-127">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>