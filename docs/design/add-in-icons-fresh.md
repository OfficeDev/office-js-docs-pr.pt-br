---
title: Diretrizes de ícone de estilo novo para os Complementos do Office
description: Diretrizes para usar ícones de estilo atualizados em Complementos do Office.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: cc891593ec9518d256047cfa172553cc41d3e12e
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604663"
---
# <a name="fresh-style-icon-guidelines-for-office-add-ins"></a>Diretrizes de ícone de estilo novo para os Complementos do Office

As versões do Office 2013+ (sem assinatura) do Office usam a iconografia de estilo Fresh da Microsoft. Se preferir que seus ícones corresponderem ao estilo Monoline do Microsoft 365, consulte Diretrizes de ícone de estilo monoline para Os [Complementos do Office](add-in-icons-monoline.md).

## <a name="office-fresh-visual-style"></a>Estilo visual do Office Fresh

Os ícones Fresh incluem apenas elementos comunicativos essenciais. Elementos não essenciais, como perspectiva, gradientes e uma fonte de luz, foram removidos. Os ícones simplificados suportam a análise mais rápida de comandos e controles. Siga esse estilo para se ajustar melhor aos clientes que não são de assinatura do Office.

## <a name="best-practices"></a>Práticas recomendadas

Siga estas diretrizes ao criar seus ícones:

|Fazer|Não fazer|
|:---|:---|
|Mantenha os elementos visuais simples e claros, concentrando-se nos principais elementos da comunicação.| Não usar artefatos que façam com que o ícone pareça confuso.|
|Usar a linguagem de ícones do Office para representar comportamentos ou conceitos.|Não reutilizar os glifos do Office UI Fabric para comandos de complemento na faixa de opções de aplicativos do Office ou menus contextuais. Os ícones do Fabric são estilisticamente diferentes e não serão compatíveis.|
|Reutilizar metáforas visuais comuns do Office, como o pincel para formatar ou a lupa para localizar.|Não reutilizar metáforas visuais para comandos diferentes. Usar o mesmo ícone para conceitos e comportamentos diferentes pode causar confusão. |
|Redesenhar os ícones para deixá-los pequenos ou maiores. Dedicar um tempo para redesenhar recortes, cantos e bordas arredondadas para maximizar a clareza da linha. |Não redimensionar os ícones reduzindo-os ou aumentando-os. Isso pode levar a uma baixa qualidade visual e a ações confusas. Os ícones complexos criados em um tamanho maior podem perder clareza ao ser redimensionados para ficar menores sem um redesenho. |
|Usar um preenchimento branco para acessibilidade. A maioria dos objetos em seus ícones exigirá um fundo branco para ser legível nos temas da interface do usuário do Office e nos modos de alto contraste.  |Evite confiar no seu logotipo ou marca para comunicar o que um comando de suplemento faz. Nem sempre é possível reconhecer as marcas em ícones menores e quando os modificadores são aplicados. Marcas de marca geralmente conflitam com estilos de ícone de faixa de opções de aplicativo do Office e podem competir pela atenção do usuário em um ambiente saturado. |
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

> [!IMPORTANT]
> Para uma imagem que seja o ícone representativo do seu complemento, consulte [Create effective listings in AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) and within Office for size and other requirements.

Não se esqueça de redesenhar seus ícones para cada tamanho em vez de reduzi-los para que caibam.

![Ilustração da recomendação para redesenhar ícones por tamanho em vez de reduzir ícones. Por exemplo, talvez seja necessário usar menos elementos em um ícone pequeno em vez de apenas dimensionar uma imagem maior.](../images/icon-resizing.png)

## <a name="icon-anatomy-and-layout"></a>Anatomia e layout do ícone

Os ícones do Office geralmente são compostos por um elemento base com modificadores conceituais e de ação sobrecarracidos. Os modificadores de ação representam conceitos como adicionar, abrir, novo ou fechar. Os modificadores conceituais representam status, alteração ou uma descrição do ícone.

Para criar comandos que se alinhem à interface do usuário do Office, siga as diretrizes de layout para o elemento básico e os modificadores. Isso garante que seus comandos tenham uma aparência profissional e seus clientes confiem no seu suplemento. Se você fizer exceções a essas diretrizes, faça intencionalmente.

A imagem a seguir mostra o layout de elementos básicos e modificadores em um ícone do Office.

![Diagrama mostrando um elemento base de ícone no centro com um modificador na parte inferior direita e um modificador de ação no canto superior esquerdo](../images/icon-layouts.png)

- Elementos básicos centrais no quadro do pixel com preenchimento todo vazio.
- Coloque modificadores de ação na parte superior esquerda.
- Coloque modificadores conceituais no canto inferior direito.
- Limite o número de elementos em seus ícones. Em 32 px, limite o número de modificadores para um máximo de dois. Em 16 px, limite o número de modificadores para um.

### <a name="base-element-padding"></a>Preenchimento do elemento básico

Coloque elementos básicos com tamanhos consistentes. Se os elementos básicos não puderem ser centralizados no quadro, alinhe-os no canto superior esquerdo, deixando os pixels extras na parte inferior direita. Para melhores resultados, aplique as diretrizes de preenchimento listadas na tabela na seção a seguir.

### <a name="modifiers"></a>Modificadores

Todos os modificadores devem ter um recorte transparente de 1 px entre cada elemento, incluindo o plano de fundo. Os elementos não devem se sobrepor diretamente. Crie um espaço em branco entre as regras e as bordas. Os modificadores podem variar um pouco de tamanho, mas use essas dimensões como ponto de partida.

|Tamanho do ícone|Preenchimento em torno do elemento básico|Tamanho do modificador|
|:---|:---|:---|
|16 px|0|9 px|
|20 px|1px|10 px|
|24 px|1px|12 px|
|32 px|2px|14 px|
|40 px|2px|20 px|
|48 px|3px|22 px|
|64 px|5px|29 px|
|80 px|5px|38 px|

## <a name="icon-colors"></a>Cores do ícone

> [!NOTE]
> Estas diretrizes de cor são destinadas a ícones da faixa de opções usados em [Comandos do suplemento](add-in-commands.md). Esses ícones não são processados com o Microsoft UI Fabric e a paleta de cores é diferente da paleta descrita em [Microsoft UI Fabric| Cores | Compartilhado](https://fluentfabric.azurewebsites.net/#/color/shared).

Os ícones do Office têm uma paleta de cores limitada. Use as cores listadas na tabela a seguir para garantir uma integração perfeita com a interface de usuário do Office. Aplique as seguintes diretrizes para o uso de cor:

- Use cor para transmitir significado, não como enfeite. Ela deve destacar ou enfatizar uma ação, status ou um elemento que diferencie explicitamente a marca.
- Se possível, use somente uma cor além do cinza. Limite as cores adicionais a no máximo duas.
- As cores devem ter uma aparência consistente em todos os tamanhos de ícone. Os ícones do Office têm paletas de cores um pouco diferentes para tamanhos de ícones diferentes. 16 px e ícones menores são ligeiramente mais escuros e mais vibrantes do que 32 px e ícones maiores. Sem esses ajustes sutis, as cores parecem variar entre os tamanhos.

|Nome da cor|RGB|Hex|Cor|Categoria|
|:---|:---|:---|:---|:---|
|Texto Cinza (80)|80, 80, 80|#505050| ![Cor cinza 80 para texto](../images/color-text-gray-80.png) |Texto|
|Texto Cinza (95)|95, 95, 95|#5F5F5F| ![Cor cinza 95 para texto](../images/color-text-gray-95.png) |Texto|
|Texto Cinza (105)|105, 105, 105|#696969| ![Cor cinza 105 para texto](../images/color-text-gray-105.png) |Texto|
|Cinza Escuro 32|128, 128, 128|#808080| ![Cor cinza escuro para 32 px e maior](../images/color-dark-gray-32.png) |32 px e acima|
|Cinza Médio 32|158, 158, 158|#9E9E9E| ![Cor cinza média para 32 px e maior](../images/color-medium-gray-32.png) |32 px e acima|
|Cinza Claro TODO|179, 179, 179|#B3B3B3| ![Cor cinza claro para todos os tamanhos de imagem](../images/color-light-gray-all.png) |Todos os tamanhos|
|Cinza Escuro 16|114, 114, 114|#727272| ![Cor cinza escuro para 16 px e menor](../images/color-dark-gray-16.png) |16 px e abaixo|
|Cinza Médio 16|144, 144, 144|#909090| ![Cor cinza média para 16 px e menor](../images/color-medium-gray-16.png) |16 e abaixo|
|Azul 32|77, 130, 184|#4d82B8| ![Cor azul para 32 px e maior](../images/color-blue-32.png) |32 px e acima|
|Azul 16|74, 125, 177|#4A7DB1| ![Cor azul para 16 px e menor](../images/color-blue-16.png) |16 px e abaixo|
|Amarelo TODO|234, 194, 130|#EAC282| ![Cor amarela para todos os tamanhos de imagem](../images/color-yellow-all.png) |Todos os tamanhos|
|Laranja 32|231, 142, 70|#E78E46| ![Cor laranja para 32 px e maior](../images/color-orange-32.png) |32 px e acima|
|Laranja 16|227, 142, 70|#E3751C| ![Cor laranja para 16 px e menor](../images/color-orange-16.png) |16 px e abaixo|
|Rosa TODO|230, 132, 151|#E68497| ![Cor rosa para todos os tamanhos de imagem](../images/color-pink-all.png) |Todos os tamanhos|
|Verde 32|118, 167, 151|#76A797| ![Cor verde para 32 px e maior](../images/color-green-32.png) |32 px e acima|
|Verde 16|104, 164, 144|#68A490| ![Cor verde para 16 px e menor](../images/color-green-16.png) |16 px e abaixo|
|Vermelho 32|216, 99, 68|#D86344| ![Cor vermelha para 32 px e maior](../images/color-red-32.png) |32 px e acima|
|Vermelho 16|214, 85, 50|#D65532| ![Cor vermelha para 16 px e menor](../images/color-red-16.png) |16 px e abaixo|
|Roxo 32|152, 104, 185|#9868B9| ![Cor roxa para 32 px e maior](../images/color-purple-32.png) |32 px e acima|
|Roxo 16|137, 89, 171|#8959AB| ![Cor roxa para 16 px e menor](../images/color-purple-16.png) |16 px e abaixo|

## <a name="icons-in-high-contrast-modes"></a>Ícones em modos de alto contraste

Os ícones do Office foram projetados para renderizar bem em modos de alto contraste. Elementos de primeiro plano são bem diferenciados dos planos de fundo para maximizar a legibilidade e habilitar a recoloração. Nos modos de alto contraste, o Office recolore qualquer pixel do seu ícone com um valor vermelho, verde ou azul menor que 190 para totalmente preto. Todos os outros pixels ficam na cor branca. Em outras palavras, cada canal RGB é avaliado onde, os valores de 0 a 189 ficam pretos e os valores de 190 a 255 ficam brancos. Outros temas de alto contraste fazem a recoloração usando o mesmo limite de valor de 190, mas com regras diferentes. Por exemplo, o tema de branco de alto contraste recolore todos pixels maiores que 190 para opaco, mas todos os outros pixels para transparente. Aplique as seguintes diretrizes para maximizar a legibilidade em configurações de alto contraste:

- Vise diferenciar elementos de primeiro plano e de plano de fundo ao longo do limite de valor de 190.
- Siga os estilos visuais dos ícones do Office.
- Use cores da nossa paleta de ícones.
- Evite o uso de gradientes.
- Evite blocos grandes de cores com valores similares.

## <a name="see-also"></a>Confira também

- [Elemento de manifesto de ícone](../reference/manifest/icon.md)
- [Elemento de manifesto IconUrl](../reference/manifest/iconurl.md)
- [Elemento de manifesto HighResolutionIconUrl](../reference/manifest/highresolutioniconurl.md)
- [Criar um ícone para o seu suplemento](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
