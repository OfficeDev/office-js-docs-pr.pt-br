---
title: Depurar suplementos no Office na Web
description: Como usar o Office na Web para testar e depurar seus suplementos.
ms.date: 03/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 58f7bfee127b69b965720ddc84c676c9f78de5bc
ms.sourcegitcommit: fb3b1c6055e664d015703623661d624251ceb6b7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/17/2022
ms.locfileid: "66136459"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Depurar suplementos no Office na Web

Este artigo descreve como usar o Office na Web para depurar seus suplementos. Use esta técnica:

- Para depurar suplementos em um computador que não esteja executando o Windows ou o cliente da área de trabalho do Office&mdash;, por exemplo, se você estiver desenvolvendo em um Mac ou Linux.
- Como um processo de depuração alternativo, se você não puder ou não desejar, depurar em um IDE, como Visual Studio ou Visual Studio Code.

Este artigo pressupõe que você tenha um projeto de suplemento que precisa ser depurado. Se você quiser apenas praticar a depuração na Web, crie um novo projeto usando um dos guias de início rápido para aplicativos Office específicos, como este início rápido para [o Word](../quickstarts/word-quickstart.md).

## <a name="debug-your-add-in"></a>Depurar o suplemento

Para depurar seu suplemento usando o Office na Web:

1. Execute o projeto no localhost e o sideload em um documento Office na Web. Para obter instruções detalhadas de [sideload, consulte Sideload Office Suplementos na Web](sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web-manually).

2. Abra as ferramentas de desenvolvedor do navegador. Isso geralmente é feito pressionando F12. Abra a ferramenta de depurador e use-a para definir pontos de interrupção e inspecionar variáveis. Para obter ajuda detalhada sobre como usar a ferramenta do navegador, consulte um dos seguintes itens.  

   - [Firefox](https://firefox-source-docs.mozilla.org/devtools-user/index.html)
   - [Safari](https://support.apple.com/guide/safari/use-the-developer-tools-in-the-develop-menu-sfri20948/mac)
   - [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
   - [Depurar suplementos usando ferramentas de desenvolvedor para Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)

   > [!NOTE]
   > Office na Web abrir no Internet Explorer.

## <a name="potential-issues"></a>Possíveis problemas

A seguir estão alguns problemas que você pode encontrar ao depurar.

- Alguns erros de JavaScript que você vê podem vir do Office na Web.

- O navegador pode mostrar um erro de certificado inválido que você deve ignorar. O processo para fazer isso varia com o navegador e as interfaces de usuário dos vários navegadores para fazer essa alteração periodicamente. Você deve pesquisar na ajuda do navegador ou pesquisar online para obter instruções. (Por exemplo, procure por "Aviso de certificado inválido do Microsoft Edge".) A maioria dos navegadores terá um link na página de aviso que permite que você clique na página do suplemento. Por exemplo, o Microsoft Edge possui um link "Ir para a página da Web (não recomendado)". Mas você geralmente terá que passar por este link toda vez que o suplemento for recarregado. Para um bypass mais duradouro, consulte a ajuda, como sugerido.

- Se você definir pontos de interrupção em seu código, Office na Web poderá gerar um erro indicando que não é possível salvar.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas para o desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)
- [Solucionar erros de usuários com Suplementos do Office](testing-and-troubleshooting.md)
