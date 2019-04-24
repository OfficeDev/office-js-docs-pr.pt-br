---
title: Diretrizes de ícone para suplementos do Office
description: ''
ms.date: 03/02/2019
localization_priority: Priority
ms.openlocfilehash: 8e741f70327584ddd1b6f51f19b276e072862229
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448839"
---
# <a name="icons"></a>Ícones
Ícones são a representação visual de um comportamento ou conceito. Eles são usados frequentemente para adicionar significado a controles e comandos. Os elementos visuais, realistas ou simbólicos, permitem ao usuário navegar pela interface do usuário da mesma maneira que os avisos os ajudam a navegar pelo ambiente. Eles devem ser simples, claros e conter apenas os detalhes necessários para permitir que os clientes analisem rapidamente a ação que ocorrerá ao escolherem um controle.

As interfaces de faixa de opções do Office têm um estilo visual padrão. Isso garante a consistência e a familiaridade em aplicativos do Office. As diretrizes ajudarão você a criar um conjunto de ativos PNG para sua solução que se ajuste como parte natural do Office.

Muitos contêineres HTML contêm controles com iconografia. Use a fonte personalizada do Office UI Fabric para renderizar os ícones com o estilo do Office no suplemento. A fonte de ícone do Fabric contém muitos glifos para metáforas comuns do Office que você pode dimensionar, colorir e estilizar para atender às suas necessidades. Se você tiver uma linguagem visual existente com seu próprio conjunto de ícones, fique à vontade para usá-la em telas HTML. Criar continuidade com sua própria marca com um conjunto de ícones padrão é uma parte importante de qualquer linguagem de design. Tenha cuidado para não criar confusão para os clientes entrando em conflito com as metáforas do Office.


## <a name="design-icons-for-add-in-commands"></a>Desenvolver ícones para comandos de suplemento

Os [Comandos de suplementos](add-in-commands.md) adicionam botões, texto e ícones à interface do usuário do Office. Os botões de comando de suplemento devem fornecer ícones significativos e rótulos que identifiquem claramente a ação que o usuário está realizando ao usar um comando. Este artigo fornece diretrizes de estilo e produção que ajudam você a desenvolver ícones que se integrem perfeitamente ao Office. 

## <a name="office-icon-design-principles"></a>Princípios de design de ícones do Office

A versão Office 2013 de clientes de área de trabalho do Office conta com uma iconografia atualizada. A mudança estilística de substituição é a redução. Os novos ícones contêm apenas elementos de comunicação essenciais. Elementos não essenciais, como perspectiva, gradientes e uma fonte de luz, foram removidos. Os ícones simplificados suportam a análise mais rápida de comandos e controles. Siga esse estilo para ter uma melhor integração com o Office.

Os ícones do Office são baseados nas seguintes princípios de design: 

- Interpretação moderna do conjunto de ícones do Office 
- Novo, porém reconhecível  
- Simples, claro e direto 

A imagem a seguir mostra ícones que aplicam os princípios modernos de design.

![Imagem mostrando ícones antigos do Office e a interpretação moderna e atualizada dos ícones](../images/icons-images.png)

## <a name="best-practices"></a>Práticas recomendadas

Siga estas diretrizes ao criar seus ícones: 

|Fazer|Não fazer|
|:---|:---|
|Manter os elementos visuais simples e claros, concentrando-se nos elementos principais da comunicação.| Não usar artefatos que façam com que o ícone pareça confuso.|
|Usar a linguagem de ícones do Office para representar comportamentos ou conceitos.|Não reutilizar glifos do Office UI Fabric para comandos de suplemento na faixa de opções do Office ou em menus contextuais. Os ícones do Fabric são estilisticamente diferentes e não serão compatíveis.|
|Reutilizar metáforas visuais comuns do Office, como o pincel para formatar ou a lupa para localizar.|Não reutilizar metáforas visuais para comandos diferentes. Usar o mesmo ícone para conceitos e comportamentos diferentes pode causar confusão. |
|Redesenhar os ícones para deixá-los pequenos ou maiores. Dedicar um tempo para redesenhar recortes, cantos e bordas arredondadas para maximizar a clareza da linha. |Não redimensionar os ícones reduzindo-os ou aumentando-os. Isso pode levar a uma baixa qualidade visual e a ações confusas. Os ícones complexos criados em um tamanho maior podem perder clareza ao ser redimensionados para ficar menores sem um redesenho. |
|Usar um preenchimento branco para acessibilidade. A maioria dos objetos em seus ícones exigirá um fundo branco para ser legível nos temas da interface do usuário do Office e nos modos de alto contraste.  ||
|Usar o formato PNG com uma tela de fundo transparente. ||
|Evitar usar conteúdo localizável nos ícones, como caracteres tipográficos, indicações de parágrafos e pontos de interrogação. ||



## <a name="icon-size-recommendations-and-requirements"></a>Recomendações e requisitos de tamanho de ícone

Os ícones da área de trabalho do Office são imagens bitmap. Os tamanhos diferentes serão renderizados, dependendo do modo de toque e da configuração de DPI do usuário. Inclua todos os oito tamanhos com suporte para criar a melhor experiência para todas as resoluções e contextos com suporte. Estes são os tamanhos compatíveis (três são obrigatórios):

- 16 px (obrigatório)
- 20 px
- 24 px
- 32 px (obrigatório)
- 40 px
- 48 px
- 64 px (recomendado, melhor para Mac)
- 80 px (obrigatório)  

Não se esqueça de redesenhar seus ícones para cada tamanho em vez de reduzi-los para que caibam.

![Ilustração que mostra a recomendação de redimensionar os ícones em vez de reduzi-los](../images/icon-resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

> [!NOTE]
> At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a>Anatomia e layout do ícone

Os ícones do Office são geralmente compostos por um elemento básico com modificadores conceituais e de ação sobrepostos. Os modificadores de ação representam conceitos como adicionar, abrir, novo ou fechar. Os modificadores conceituais representam status, alteração ou uma descrição do ícone. 

Para criar comandos que se alinhem à interface do usuário do Office, siga as diretrizes de layout para o elemento básico e os modificadores. Isso garante que seus comandos tenham uma aparência profissional e seus clientes confiem no seu suplemento. Se você fizer exceções a essas diretrizes, faça intencionalmente.

A imagem a seguir mostra o layout de elementos básicos e modificadores em um ícone do Office.

![Imagem mostrando o elemento básico de um ícone no centro com um modificador no canto inferior direito e um modificador de ação no canto superior esquerdo](../images/icon-layouts.png)

- Elementos básicos centrais no quadro do pixel com preenchimento todo vazio.
- Coloque modificadores de ação na parte superior esquerda. 
- Coloque modificadores conceituais no canto inferior direito.
- Limite o número de elementos em seus ícones. Com 32 px, limite o número de modificadores a no máximo dois. Com 16 px, limite o número de modificadores a um.

###<a name="base-element-padding"></a>Preenchimento do elemento básico
Coloque elementos básicos com tamanhos consistentes. Se os elementos básicos não puderem ser centralizados no quadro, alinhe-os no canto superior esquerdo, deixando os pixels extras na parte inferior direita. Para melhores resultados, aplique as diretrizes de preenchimento listadas na tabela a seguir.

###<a name="modifiers"></a>Modificadores
Todos os modificadores devem ter um recorte transparente de 1 px entre cada elemento, incluindo a tela de fundo. Os elementos não devem se sobrepor diretamente. Crie um espaço em branco entre as regras e as bordas. Os modificadores podem variar um pouco de tamanho, mas use essas dimensões como ponto de partida.


|**Tamanho do ícone**|**Preenchimento em torno do elemento básico**|**Tamanho do modificador**|
|:---|:---|:---|
|16px|0|9px|
|20px|1px|10px|
|24px|1px|12px|
|32px|2px|14px|
|40px|2px|20px|
|48px|3px|22px|
|64px|5px|29px|
|80px|5px|38px|


## <a name="icon-colors"></a>Cores do ícone

> [!NOTE]
> Estas diretrizes de cor são destinadas a ícones da faixa de opções usados em [Comandos do suplemento](add-in-commands.md). Esses ícones não são processados com o Microsoft UI Fabric e a paleta de cores é diferente da paleta descrita em [Microsoft UI Fabric| Cores | Compartilhado](https://fluentfabric.azurewebsites.net/#/color/shared).

Os ícones do Office têm uma paleta de cores limitada. Use as cores listadas na tabela a seguir para garantir uma integração perfeita com a interface de usuário do Office. Aplique as seguintes diretrizes para o uso de cor: 

- Use cor para transmitir significado, não como enfeite. Ela deve destacar ou enfatizar uma ação, status ou um elemento que diferencie explicitamente a marca.  
- Se possível, use somente uma cor além do cinza. Limite as cores adicionais a no máximo duas.
- As cores devem ter uma aparência consistente em todos os tamanhos de ícone. Os ícones do Office têm paletas de cores um pouco diferentes para tamanhos de ícones diferentes. Ícones com 16 px e menores são um pouco mais escuros e mais vibrantes do que os ícones de 32 px e maiores. Sem esses ajustes sutis, as cores parecem variar entre os tamanhos.   

|**Nome da cor**|**RGB**|**Hex**|**Cor**|**Categoria**|
|:---|:---|:---|:---|:---|
|Texto Cinza (80)|80, 80, 80|#505050| ![Imagem colorida texto cinza 80](../images/color-text-gray-80.png) |Texto|
|Texto Cinza (95)|95, 95, 95|#5F5F5F| ![Imagem colorida texto cinza 95](../images/color-text-gray-95.png) |Texto|
|Texto Cinza (105)|105, 105, 105|#696969| ![Imagem colorida texto cinza 105](../images/color-text-gray-105.png) |Texto|
|Cinza Escuro 32|128, 128, 128|#808080| ![Imagem colorida cinza escuro 32](../images/color-dark-gray-32.png) |32 e acima|
|Cinza Médio 32|158, 158, 158|#9E9E9E| ![Imagem colorida cinza médio 32](../images/color-medium-gray-32.png) |32 e acima|
|Cinza Claro TODO|179, 179, 179|#B3B3B3| ![Imagem colorida cinza claro todo](../images/color-light-gray-all.png) |Todos os tamanhos|
|Cinza Escuro 16|114, 114, 114|#727272| ![Imagem colorida cinza escuro 16](../images/color-dark-gray-16.png) |16 e abaixo|
|Cinza Médio 16|144, 144, 144|#909090| ![Imagem colorida cinza médio 16](../images/color-medium-gray-16.png) |16 e abaixo|
|Azul 32|77, 130, 184|#4d82B8| ![Imagem colorida azul 32](../images/color-blue-32.png) |32 e acima|
|Azul 16|74, 125, 177|#4A7DB1| ![Imagem colorida azul 16](../images/color-blue-16.png) |16 e abaixo|
|Amarelo TODO|234, 194, 130|#EAC282| ![Imagem colorida amarelo todo](../images/color-yellow-all.png) |Todos os tamanhos|
|Laranja 32|231, 142, 70|#E78E46| ![Imagem colorida laranja 32](../images/color-orange-32.png) |32 e acima|
|Laranja 16|227, 142, 70|#E3751C| ![Imagem colorida laranja 16](../images/color-orange-16.png) |16 e abaixo|
|Rosa TODO|230, 132, 151|#E68497| ![Imagem colorida rosa todo](../images/color-pink-all.png) |Todos os tamanhos|
|Verde 32|118, 167, 151|#76A797| ![Imagem colorida verde 32](../images/color-green-32.png) |32 e acima|
|Verde 16|104, 164, 144|#68A490| ![Imagem colorida verde 16](../images/color-green-16.png) |16 e abaixo|
|Vermelho 32|216, 99, 68|#D86344| ![Imagem colorida vermelho 32](../images/color-red-32.png) |32 e acima|
|Vermelho 16|214, 85, 50|#D65532| ![Imagem colorida vermelho 16](../images/color-red-16.png) |16 e abaixo|
|Roxo 32|152, 104, 185|#9868B9| ![Imagem colorida roxo 32](../images/color-purple-32.png) |32 e acima|
|Roxo 16|137, 89, 171|#8959AB| ![Imagem colorida roxo 16](../images/color-purple-16.png) |16 e abaixo|


## <a name="icons-in-high-contrast-modes"></a>Ícones em modos de alto contraste

Os ícones do Office foram projetados para renderizar bem em modos de alto contraste. Elementos de primeiro plano são bem diferenciados dos planos de fundo para maximizar a legibilidade e habilitar a recoloração. Nos modos de alto contraste, o Office recolore qualquer pixel do seu ícone com um valor vermelho, verde ou azul menor que 190 para totalmente preto. Todos os outros pixels ficam na cor branca. Em outras palavras, cada canal RGB é avaliado onde, os valores de 0 a 189 ficam pretos e os valores de 190 a 255 ficam brancos. Outros temas de alto contraste fazem a recoloração usando o mesmo limite de valor de 190, mas com regras diferentes. Por exemplo, o tema de branco de alto contraste recolore todos pixels maiores que 190 para opaco, mas todos os outros pixels para transparente. Aplique as seguintes diretrizes para maximizar a legibilidade em configurações de alto contraste:

- Vise diferenciar elementos de primeiro plano e de plano de fundo ao longo do limite de valor de 190.
- Siga os estilos visuais dos ícones do Office.
- Use cores da nossa paleta de ícones.
- Evite o uso de gradientes.
- Evite blocos grandes de cores com valores similares.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas de desenvolvimento de suplementos](../concepts/add-in-development-best-practices.md)
- [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md)




- Evite confiar no seu logotipo ou marca para comunicar o que um comando de suplemento faz. Nem sempre é possível reconhecer as marcas em ícones menores e quando os modificadores são aplicados. As marcas geralmente entram em conflito com estilos de ícone da faixa de opções e podem competir pela atenção do usuário em um ambiente saturado.


