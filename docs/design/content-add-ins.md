---
title: Suplementos de conteúdo do Office
description: Suplementos de conteúdo são superfícies que podem ser incorporadas diretamente em documentos do Excel ou do PowerPoint que concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f228ae8e7cca0426b0b43e31e38454029e4c7614
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093844"
---
# <a name="content-office-add-ins"></a>Suplementos de conteúdo do Office

Suplementos de conteúdo são superfícies que podem ser incorporadas diretamente em documentos do Excel ou PowerPoint. Os suplementos de conteúdo concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento.  

*Figura 1. Layout típico dos suplementos de conteúdo*

![Imagem de exemplo exibindo um layout típico de suplementos de conteúdo.](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Práticas recomendadas

- Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.
- Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento (aplica-se apenas a suplementos do Excel e do PowerPoint).

## <a name="variants"></a>Variantes

Tamanhos de suplementos de conteúdo para Excel e PowerPoint na área de trabalho do Office e o Microsoft 365 são especificados pelo usuário.

## <a name="personality-menu"></a>Menu de personalidade

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.

*Figura 2. Menu de personalidade no Windows* 

![Imagem mostrando o menu do personalidade na área de trabalho do Windows](../images/personality-menu-win.png)


No Mac, o menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço ocupado para 34 x 32 pixels, como mostrado.

*Figura 3. Menu de personalidade no Mac*

![Imagem mostrando o menu de personalidade na área de trabalho do Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implementação

Para ver um exemplo que implementa um suplemento de conteúdo, confira [Suplemento de conteúdo do Excel Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) no GitHub.

## <a name="support-considerations"></a>Considerações sobre o suporte

- Verifique se os suplementos do Office funcionarão em uma [plataforma de host do Office específica](../overview/office-add-in-availability.md).
- Alguns suplementos de conteúdo podem obrigar o usuário a "confiar" nele para ler e gravar no Excel ou PowerPoint. Você pode declarar no manifesto do suplemento quais [níveis de permissão](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) deseja que o usuário tenha.  
- Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later. If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.

## <a name="see-also"></a>Confira também

- [Disponibilidade de host e plataforma para suplementos do Office](../overview/office-add-in-availability.md)
- [Office UI Fabric em Suplementos do Office](../design/office-ui-fabric.md)
- [Padrões de design da experiência do usuário para suplementos do Office](../design/ux-design-pattern-templates.md)
- [Solicitar permissões para uso da API em suplementos ](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
