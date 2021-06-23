---
title: Realizar o sideload de suplementos do Office em um iPad ou Mac para teste
description: Teste seu Office de iPad e Mac ao fazer sideload.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: b3d7d7fa3ee69e849c112c888b66fa9deed23d88
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076200"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>Realizar o sideload de suplementos do Office em um iPad ou Mac para teste

Para ver como seu suplemento será executado no Office no iOS, você pode realizar o sideload do manifesto do seu suplemento em um iPad usando o iTunes ou realizar o sideload do manifesto do suplemento diretamente no Office no Mac. Esta ação não permite definir pontos de interrupção e depurar o código do seu suplemento enquanto ele estiver em execução, mas é possível ver como ele se comporta e verificar se a interface do usuário é utilizável e está sendo processada adequadamente.

## <a name="prerequisites-for-office-on-ios"></a>Pré-requisitos do Office no iOS

- Um computador com Windows ou Mac com [iTunes](https://www.apple.com/itunes/download/) instalado.
  > [!IMPORTANT]
  > Se você estiver executando o macOS Catalina, [o iTunes](https://support.apple.com/HT210200) não estará mais disponível, portanto, você deve seguir as instruções na seção Sideload de um complemento no Excel ou word no iPad usando [macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) mais adiante neste artigo.

- Um iPad executando o iOS 8.2 ou posterior com Excel [ou](https://apps.apple.com/app/microsoft-excel/id586683407) [Word](https://apps.apple.com/app/microsoft-word/id586447913) instalado e um cabo de sincronização.

- O arquivo de manifesto .xml para o suplemento que você deseja testar.

## <a name="prerequisites-for-office-on-mac"></a>Pré-requisitos do Office no Mac

- Um Mac executando OS X v10.10 “Yosemite” ou posterior com [Office no Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) instalado.

- Word no Mac versão 15.18 (160109).

- Excel no Mac versão 15.19 (160206).

- PowerPoint no Mac versão 15.24 (160614)

- O arquivo de manifesto .xml para o suplemento que você deseja testar.

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>Fazer sideload de um complemento no Excel ou no Word no iPad usando o iTunes

1. Use um cabo de sincronização para conectar seu iPad ao computador. Se você estiver conectando o iPad ao computador pela primeira vez, será solicitado a confiar **neste computador?**. Escolha **Confiar** para continuar.

2. No iTunes, escolha o **ícone iPad** abaixo da barra de menus.

3. Em **Configurações** lado esquerdo do iTunes, escolha **Aplicativos**.

4. No lado direito do iTunes, role para baixo  até **Compartilhamento** de Arquivos e escolha Excel ou **Word** na coluna **Complementos.**

5. Na parte inferior da coluna **Excel** ou Documentos do **Word,** escolha **Adicionar** Arquivo e selecione o arquivo .xml de manifesto do complemento que você deseja fazer sideload.

6. Abra o aplicativo Excel ou Word em seu iPad. Se o Excel ou o aplicativo word já estiver em execução, escolha o botão **Início** e feche e reinicie o aplicativo.

7. Abra um documento.

8. Escolha **Complementos na**  guia Inserir. (Na guia Inserir, talvez seja necessário rolar horizontalmente até ver o botão **Adicionar.)**  Seu complemento sideload está disponível para ser inserido no título **Desenvolvedor** na interface do usuário **de complementos.**

    ![Insira os complementos no aplicativo Excel.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>Fazer sideload de um complemento no Excel ou no Word no iPad usando macOS Catalina

> [!IMPORTANT]
> Com a introdução do macOS Catalina, a Apple descontinuou [o iTunes](https://support.apple.com/HT210200) no Mac e a funcionalidade integrada necessária para sideload de aplicativos **no Finder**.

1. Use um cabo de sincronização para conectar seu iPad ao computador. Se você estiver conectando o iPad ao computador pela primeira vez, será solicitado a confiar **neste computador?**. Escolha **Confiar** para continuar. Você também pode ser perguntado se esse é um novo iPad ou se você está restaurando um.

2. No Localizador, em **Locais,** escolha **o** ícone iPad abaixo da barra de menus.

3. Na parte superior da janela Localizador, clique em **Arquivos** e, em seguida, localize **Excel** ou **Word**.

4. Em uma janela do Finder diferente, arraste e solte o arquivo manifest.xml do complemento que você deseja carregar no arquivo **Excel** ou **Word** na primeira janela do Finder.

5. Abra o aplicativo Excel ou Word em seu iPad. Se o Excel ou o aplicativo word já estiver em execução, escolha o botão **Início** e feche e reinicie o aplicativo.

6. Abra um documento.

7. Escolha **Complementos na**  guia Inserir. (Na guia Inserir, talvez seja necessário rolar horizontalmente até ver o botão **Adicionar.)**  Seu complemento sideload está disponível para ser inserido no título **Desenvolvedor** na interface do usuário **de complementos.**

    ![Insira os complementos no aplicativo Excel.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a>Realizar sideload de um suplemento no Office no Mac

> [!NOTE]
> Para realizar o sideload de um suplemento do Outlook no Mac, confira [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md).

1. Abra **Terminal** e vá para uma das seguintes pastas onde você salvará o arquivo de manifesto do seu complemento. Se a pasta `wef` não existir em seu computador, crie-a.

    - Para o Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Para o Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - Para o PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

2. Abra a pasta no **Finder** usando o comando `open .` (incluindo o ponto ou ponto). Copie o arquivo de manifesto do suplemento nessa pasta.

    ![Pasta Wef em Office no Mac.](../images/all-my-files.png)

3. Abra o Word e abra um documento. Reinicie o Word se já estiver em execução.

4. No Word, **escolha Inserir** Meus  >    >  **Complementos** (menu suspenso) e escolha o seu complemento.

    ![Meus Complementos no Office no Mac.](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Aplicativos em que foi feito o sideload não aparecerão na caixa de diálogo Meus Suplementos. Eles só ficam visíveis dentro do menu suspenso (pequena seta para baixo à direita de Meus Suplementos na guia **Inserir**). Os suplementos em que foi feito o sideload são exibidos na lista sob o título **Suplementos do Desenvolvedor** nesse menu.

5. Verifique se o seu suplemento é exibido no Word.

    ![Office O complemento exibido no Office no Mac.](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>Remover um complemento com sideload

Você pode remover um complemento com sideload anteriormente limpando o cache Office em seu computador. Detalhes sobre como limpar o cache de cada plataforma e aplicativo podem ser encontrados no artigo [Limpar o Office cache](clear-cache.md).

## <a name="see-also"></a>Confira também

- [Depurar suplementos do Office no iPad e no Mac](debug-office-add-ins-on-ipad-and-mac.md)
