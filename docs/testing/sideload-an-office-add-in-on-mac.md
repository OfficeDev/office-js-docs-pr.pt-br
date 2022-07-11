---
title: Realizar sideload de Suplementos do Office no Mac para teste
description: Teste seu Suplemento do Office no Mac ao realizar o sideload.
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38ed5f5dba2d379b6137a098240021bd642d6e11
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713197"
---
# <a name="sideload-office-add-ins-on-mac-for-testing"></a>Realizar sideload de Suplementos do Office no Mac para teste

Para ver como o suplemento será executado no Office no Mac, você pode fazer o sideload do manifesto do suplemento. Esta ação não permite definir pontos de interrupção e depurar o código do seu suplemento enquanto ele estiver em execução, mas é possível ver como ele se comporta e verificar se a interface do usuário é utilizável e está sendo processada adequadamente.

> [!NOTE]
> Para realizar o sideload de um suplemento do Outlook, confira [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md).

## <a name="prerequisites-for-office-on-mac"></a>Pré-requisitos do Office no Mac

- Um Mac executando OS X v10.10 “Yosemite” ou posterior com [Office no Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) instalado.

- Word no Mac versão 15.18 (160109).

- Excel no Mac versão 15.19 (160206).

- PowerPoint no Mac versão 15.24 (160614).

- O arquivo de manifesto .xml para o suplemento que você deseja testar.

## <a name="sideload-an-add-in-in-office-on-mac"></a>Realizar sideload de um suplemento no Office no Mac

1. Use **o Finder** para fazer sideload do arquivo de manifesto. Abra **o Finder** e insira Command+Shift+G para abrir a caixa **de diálogo Ir para pasta** .

1. Insira um dos seguintes caminhos de arquivo, com base no aplicativo que você deseja usar para sideload. Se a pasta `wef` não existir em seu computador, crie-a.

    - Para o Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Para o Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - Para o PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

        > [!NOTE]
        > As etapas restantes descrevem como fazer sideload de um suplemento do Word.

1. Copie o arquivo de manifesto do suplemento para essa `wef` pasta.

    ![Pasta Wef no Office no Mac.](../images/all-my-files.png)

1. Abra o Word e abra um documento. Reinicie o Word se já estiver em execução.

1. No Word, **escolha Inserir** > **Suplementos** >  Meus **Suplementos** (menu suspenso) e escolha seu suplemento.

    ![Meus Suplementos no Office no Mac.](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Aplicativos em que foi feito o sideload não aparecerão na caixa de diálogo Meus Suplementos. Eles só ficam visíveis dentro do menu suspenso (pequena seta para baixo à direita de Meus Suplementos na guia **Inserir**). Os suplementos em que foi feito o sideload são exibidos na lista sob o título **Suplementos do Desenvolvedor** nesse menu.

1. Verifique se o seu suplemento é exibido no Word.

    ![Suplemento do Office exibido no Office no Mac.](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>Remover um suplemento de sideload

Você pode remover um suplemento com sideload anteriormente limpando o cache do Office em seu computador. Detalhes sobre como limpar o cache para cada plataforma e aplicativo podem ser encontrados no artigo [Limpar o cache do Office](clear-cache.md).

## <a name="see-also"></a>Confira também

- [Realizar sideload de Suplementos do Office no iPad para teste](sideload-an-office-add-in-on-ipad.md)
- [Depurar Suplementos do Office em um Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md)
