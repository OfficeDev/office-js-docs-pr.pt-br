---
title: Diretrizes de tipografia para suplementos do Office
description: Saiba o que são os tipos de fonte e de fontes a serem usados nos suplementos do Office.
ms.date: 06/27/2018
localization_priority: Normal
ms.openlocfilehash: 9f9398137c9e8a00a2743e99d94e405ff80d85dd
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607656"
---
# <a name="typography"></a>Tipografia

Segoe é o tipo de fonte padrão para o Office. Use-a no suplemento para alinhar objetos de conteúdo, caixas de diálogo e painéis de tarefas do Office. O Office UI Fabric lhe dá acesso à fonte Segoe. Ele fornece um conjunto completo da fonte Segoe com muitas variações (incluindo espessura e tamanho da fonte) em classes CSS convenientes. Nem todos os tamanhos e espessuras do Office UI Fabric terão boa aparência em um suplemento do Office. Para obter um ajuste harmonioso ou evitar conflitos, considere o uso de um subconjunto do conjunto de fontes do Fabric. Aqui está uma lista de classes base do Fabric que recomendamos para uso em suplementos do Office.

|Amostra |Classe |Tamanho |Peso |Uso recomendado |
|------ |----- |---- |------ |----------------- |
|![Imagem de texto Hero](../images/add-in-typeramp-hero.png)|.ms-font-xxl |28 px | Segoe Light |<ul><li>Essa classe é maior do que todos os outros elementos tipográficos no Office. Use-a com moderação para não prejudicar o ajuste na hierarquia visual.</li><li>Evite o uso de cadeias de caracteres longas em espaços restritos.</li><li>Deixe bastante espaço em branco ao redor do texto ao usar esta classe.</li><li>Comumente usada para mensagens da tela de apresentação, elementos Hero ou outras chamadas à ação.</li></ul> |
|![Imagem de texto Hero](../images/add-in-typeramp-title.png)|.ms-font-xl |21 px |Segoe Light | <ul><li>Essa classe corresponde ao título do painel de tarefas dos aplicativos do Office.</li><li>Use-a com moderação para evitar uma hierarquia tipográfica monótona.</li><li>Comumente usado como o elemento de nível superior, como títulos de conteúdo, página ou caixa de diálogo.</li></ul> |
|![Imagem de texto Hero](../images/add-in-typeramp-subtitle.png)|.ms-font-l |17 px |Segoe Semilight | <ul><li>Essa classe é a primeira abaixo de títulos.</li><li>Comumente usada como um subtítulo, um elemento de navegação ou um cabeçalho de grupo.</li><ul> |
|![Imagem de Texto Hero](../images/add-in-typeramp-body.png)|.ms-font-m |14 px |Segoe Regular |<ul><li>Muito usada como corpo de texto dentro de suplementos.</li><ul>|
|![Imagem de texto Hero](../images/add-in-typeramp-caption.png)|.ms-font-xs |11 px | Segoe Regular |<ul><li>Muito usada em texto secundário ou terciário, como carimbos de data/hora, linhas, títulos ou rótulos de campo.</li><ul>|
|![Imagem de texto Hero](../images/add-in-typeramp-annotation.png)|.ms-font-mi |10 px |Segoe Semibold |<ul><li>A menor etapa no painel de tipos deve ser usada raramente. Está disponível para situações em que a legibilidade não é necessária.</li><ul>|

> [!NOTE]
> A cor do texto não está incluída nessas classes base. Use a opção “Neutro principal” do Fabric para a maioria dos textos em fundos brancos.
