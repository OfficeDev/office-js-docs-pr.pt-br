---
title: Diretrizes de ícone de estilo monolinha para Office de complementos
description: Diretrizes para usar ícones de estilo monoline em Office de complementos.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 0e8bf4f39ddbad457df7d033a08836825d9e1d3f
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938109"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a>Diretrizes de ícone de estilo monolinha para Office de complementos

Iconografia de estilo monoline é usada em Office aplicativos. Se você preferir que seus ícones corresponderem ao estilo Fresh de não assinatura Office 2013+, consulte Diretrizes de ícone de estilo novo [para Office Add-ins](add-in-icons-fresh.md).

## <a name="office-monoline-visual-style"></a>Office Estilo visual monoline

O objetivo do estilo Monoline é ter uma iconografia consistente, clara e acessível para comunicar ações e recursos com elementos visuais simples, garantir que os ícones sejam acessíveis a todos os usuários e tenham um estilo consistente com os usados em outros locais Windows.

As diretrizes a seguir são para desenvolvedores de terceiros que querem criar ícones para recursos que serão consistentes com os ícones já presentes Office produtos.

### <a name="design-principles"></a>Princípios de design

- Simples, limpo, claro.
- Contém apenas elementos necessários.
- Inspirado no estilo Windows ícone.
- Acessível a todos os usuários.

#### <a name="convey-meaning"></a>Transmitir significado

- Use elementos descritivos, como uma página, para representar um documento ou um envelope para representar o email.
- Use o mesmo elemento para representar o mesmo conceito, ou seja, o email é sempre representado por um envelope, não por um carimbo.
- Use uma metáfora principal durante o desenvolvimento de conceitos.

#### <a name="reduction-of-elements"></a>Redução de elementos

- Reduza o ícone ao seu significado principal, usando apenas elementos essenciais à metáfora.
- Limite o número de elementos em um ícone para dois, independentemente do tamanho do ícone.

#### <a name="consistency"></a>Consistência

Tamanhos, disposição e cor dos ícones devem ser consistentes.

#### <a name="styling"></a>Estilo

##### <a name="perspective"></a>Perspectiva

Os ícones monoline são voltados para frente por padrão. Certos elementos que exigem perspectiva e/ou rotação, como um cubo, são permitidos, mas as exceções devem ser mantidas no mínimo.

##### <a name="embellishment"></a>Embelezamento

Monoline é um estilo mínimo limpo. Tudo usa cor plana, o que significa que não há gradientes, texturas ou fontes de luz.

## <a name="designing"></a>Designing

### <a name="sizes"></a>Tamanhos

Recomendamos que você produza cada ícone em todos esses tamanhos para dar suporte a dispositivos DPI altos. Os *tamanhos absolutamente* necessários são 16 px, 20 px e 32 px, pois esses são os tamanhos 100%.

**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**

> [!IMPORTANT]
> Para uma imagem que é o ícone representativo do seu complemento, consulte [Create effective listings in AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) and within Office for size and other requirements.

### <a name="layout"></a>Layout

A seguir, um exemplo de layout de ícone com um modificador.

![Diagrama do ícone com modificador no canto inferior direito.](../images/monolineicon1.png)  ![Diagrama do mesmo ícone com plano de fundo de grade adicionado e textos explicadores para a base, modificador, preenchimento e recorte.](../images/monolineicon2.png)

#### <a name="elements"></a>Elementos

- **Base**: O conceito principal que o ícone representa. Normalmente, esse é o único visual necessário para o ícone, mas às vezes o conceito principal pode ser aprimorado com um elemento secundário, um modificador.

- **Modificador** Qualquer elemento que sobrepõe a base; ou seja, um modificador que normalmente representa uma ação ou um status. Modifica o elemento base agindo como uma adição, alteração ou descritor.

![Diagrama de grade com áreas base e modificadora chamadas.](../images/monolineicon3.png)

### <a name="construction"></a>Construção

#### <a name="element-placement"></a>Posicionamento do elemento

Os elementos base são colocados no centro do ícone dentro do preenchimento. Se não puder ser colocado perfeitamente centralizado, a base deve errá-la para a direita superior. No exemplo a seguir, o ícone é perfeitamente centralizado.

![Diagrama mostrando ícone perfeitamente centralizado.](../images/monolineicon4.png)

No exemplo a seguir, o ícone está errando para a esquerda.

![Diagrama mostrando o ícone que erra para a esquerda por 1 px.](../images/monolineicon5.png)

Os modificadores quase sempre são colocados no canto inferior direito da tela de ícone. Em alguns casos raros, os modificadores são colocados em um canto diferente. Por exemplo, se o elemento base não for reconhecido com o modificador no canto inferior direito, considere colocá-lo no canto superior esquerdo.

![Diagrama mostrando quatro ícones com o modificador na parte inferior direita e um ícone com o modificador no canto superior esquerdo.](../images/monolineicon6.png)

#### <a name="padding"></a>Padding

Cada ícone de tamanho tem uma quantidade especificada de preenchimento ao redor do ícone. O elemento base permanece dentro do preenchimento, mas o modificador deve ficar até a borda da tela, estendendo-se fora do preenchimento até a borda da borda do ícone. As imagens a seguir mostram o preenchimento recomendado a ser usado para cada um dos tamanhos de ícone.

|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![Ícone de 16 px com preenchimento 0px.](../images/monolineicon7.png)|![Ícone de 20 px com preenchimento de 1px.](../images/monolineicon8.png)|![Ícone de 24 px com preenchimento de 1px.](../images/monolineicon9.png)|![Ícone de 32 px com preenchimento de 2px.](../images/monolineicon10.png)|![Ícone de 40 px com preenchimento de 2px.](../images/monolineicon11.png)|![Ícone de 48 px com preenchimento de 3px.](../images/monolineicon12.png)|![Ícone de 64 px com preenchimento de 4px.](../images/monolineicon13.png)|![Ícone de 80 px com preenchimento de 5px.](../images/monolineicon14.png)|![Ícone px de 96 com preenchimento de 6px.](../images/monolineicon15.png)|

#### <a name="line-weights"></a>Pesos de linha

Monoline é um estilo dominado por formas de linha e delineadas. Dependendo do tamanho que você está produzindo, o ícone deve usar os seguintes pesos de linha.

|Tamanho do ícone:|16px|20px|24px|32px|40px|48px|64px|80px|96px|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**Peso da linha:**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
|**Ícone de exemplo:**|![Ícone de 16 px.](../images/monolineicon16.png)|![Ícone de 20 px.](../images/monolineicon17.png)|![Ícone de 24 px.](../images/monolineicon18.png)|![Ícone de 32 px.](../images/monolineicon19.png)|![Ícone de 40 px.](../images/monolineicon20.png)|![Ícone de 48 px.](../images/monolineicon21.png)|![Ícone de 64 px.](../images/monolineicon22.png)|![Ícone de 80 px.](../images/monolineicon23.png)|![Ícone px de 96.](../images/monolineicon24.png)|

#### <a name="cutouts"></a>Recortes

Quando um elemento icon é colocado sobre outro elemento, um recorte (do elemento inferior) é usado para fornecer espaço entre os dois elementos, principalmente para fins de leitura. Isso geralmente acontece quando um modificador é colocado sobre um elemento base, mas também há casos em que nenhum dos elementos é um modificador. Esses recortes entre os dois elementos são às vezes chamados de "lacuna".

O tamanho da lacuna deve ter a mesma largura que o peso da linha usado nesse tamanho. Se você criar um ícone de 16 px, a largura da lacuna será 1px e se for um ícone de 48 px, o intervalo deverá ser de 2px. O exemplo a seguir mostra um ícone de 32 px com um intervalo de 1px entre o modificador e a base subjacente.

![Ícone de 32 px com um intervalo de 1px entre o modificador e a base subjacente.](../images/monolineicon25.png)

Em alguns casos, o intervalo pode ser maior em um px de 1/2 se o modificador tiver uma borda diagonal ou curva e o intervalo padrão não fornecer separação suficiente. Isso provavelmente afetará apenas os ícones com peso de linha de 1px: 16 px, 20 px, 24 px e 32 px.

#### <a name="background-fills"></a>Preenchimentos em segundo plano

A maioria dos ícones no conjunto de ícones monoline exige preenchimentos em segundo plano. No entanto, há casos em que o objeto não teria naturalmente um preenchimento, portanto, nenhum preenchimento deve ser aplicado. Os ícones a seguir têm um preenchimento em branco.

![Compilação de cinco ícones com preenchimento branco.](../images/monolineicon26.png)

Os ícones a seguir não têm preenchimento. (O ícone de engrenagem está incluído para mostrar que o buraco central não está preenchido.)

![Compilação de cinco ícones sem preenchimento.](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a>Práticas recomendadas para preenchimentos

###### <a name="dos"></a>O que fazer

- Preencha qualquer elemento que tenha um limite definido e que tenha um preenchimento naturalmente.
- Use uma forma separada para criar o preenchimento em segundo plano.
- Use **o Preenchimento de** Plano de Fundo da [paleta de cores](#color).
- Mantenha a separação de pixels entre elementos sobrepostos.
- Preencha entre vários objetos.

###### <a name="donts"></a>O que não fazer

- Não preencha objetos que não seriam preenchidos naturalmente; por exemplo, um clipe de papel.
- Não preencha colchetes.
- Não preencha por trás de números ou caracteres alfa.

### <a name="color"></a>Cor

A paleta de cores foi projetada para simplicidade e acessibilidade. Ele contém 4 cores neutras e duas variações para azul, verde, amarelo, vermelho e roxo. Laranja não está incluída intencionalmente na paleta de cores do ícone monoline. Cada cor destina-se a ser usada de maneiras específicas, conforme descrito nesta seção.

#### <a name="palette"></a>Paleta

![Os quatro tons de cinza em monoline: cinza escuro para autônomo ou contorno, cinza médio para contorno ou conteúdo, cinza muito claro para preenchimento de plano de fundo e cinza claro para preenchimento.](../images/monoline-grayshades.png)

![A paleta de cores em monoline inclui um tom de azul, verde, amarelo, vermelho e roxo para autônomo, contorno e preenchimento.](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a>Como usar cor

Na paleta de cores Monoline, todas as cores têm variações Autônomas, Contornos e Preenchimento. Geralmente, os elementos são construídos com um preenchimento e uma borda. As cores são aplicadas em um dos padrões a seguir.

- A cor autônoma sozinha para objetos que não têm preenchimento.
- A borda usa a cor Outline e o preenchimento usa a cor Preenchimento.
- A borda usa a cor Autônoma e o preenchimento usa a cor Preenchimento de Plano de Fundo.

A seguir estão exemplos de uso de cor.

![Compilação de três ícones com cor em uma borda ou preenchimento ou ambos.](../images/monolineicon28.png)

A situação mais comum será ter um elemento que use Dark Gray Standalone with Background Fill.

Ao usar um Preenchimento colorido, ele sempre deve estar com sua cor Delineada correspondente. Por exemplo, Preenchimento Azul só deve ser usado com o Contorno Azul. Mas há duas exceções para essa regra geral.

- O Preenchimento de Plano de Fundo pode ser usado com qualquer cor Autônoma.
- O Preenchimento Cinza Claro pode ser usado com duas cores de outline diferentes: Cinza Escuro ou Cinza Médio.

#### <a name="when-to-use-color"></a>Quando usar cor

A cor deve ser usada para transmitir o significado do ícone em vez de para embelezamento. Ele deve **realçar a ação** para o usuário. Quando um modificador é adicionado a um elemento base que tem cor, o elemento base normalmente é transformado em Cinza Escuro e Preenchimento de Plano de Fundo para que o modificador possa ser o elemento de cor, como o caso abaixo com o modificador "X" sendo adicionado à base de imagem no ícone mais à esquerda do conjunto a seguir.

![Compilação de cinco ícones que usam cor.](../images/monolineicon29.png)

Você deve limitar seus ícones **a uma** cor adicional, diferente das opções Outline e Fill mencionadas acima. No entanto, mais cores podem ser usadas se for vital para sua metáfora, com um limite de duas cores adicionais que não sejam cinza. Em casos raros, há exceções quando mais cores são necessárias. A seguir estão bons exemplos de ícones que usam apenas uma cor.

  ![Compilação de cinco ícones que cada um usa uma cor.](../images/monolineicon30.png)

Mas os ícones a seguir usam muitas cores.

  ![Compilação de cinco ícones que cada um usa várias cores.](../images/monolineicon31.png)

Use **Cinza Médio** para "conteúdo" interno, como linhas de grade em um ícone de uma planilha. Cores internas adicionais são usadas quando o conteúdo precisa mostrar o comportamento do controle.

![Compilação de cinco ícones com elementos internos cinza médios.](../images/monolineicon32.png)

#### <a name="text-lines"></a>Linhas de texto

Quando as linhas de texto estão em um "contêiner" (por exemplo, texto em um documento), use cinza médio. Linhas de texto que não estão em um contêiner devem ser **Cinza Escuro**.

### <a name="text"></a>Texto

Evite usar caracteres de texto em ícones. Como Office produtos são usados em todo o mundo, queremos manter os ícones o mais neutro possível.

## <a name="production"></a>Produção

### <a name="icon-file-format"></a>Formato de arquivo icon

Os ícones finais devem ser salvos como arquivos .png imagem. Use o formato PNG com um plano de fundo transparente e tenha profundidade de 32 bits.

## <a name="see-also"></a>Confira também

- [Elemento de manifesto de ícone](../reference/manifest/icon.md)
- [Elemento de manifesto IconUrl](../reference/manifest/iconurl.md)
- [Elemento de manifesto HighResolutionIconUrl](../reference/manifest/highresolutioniconurl.md)
- [Criar um ícone para o seu suplemento](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
