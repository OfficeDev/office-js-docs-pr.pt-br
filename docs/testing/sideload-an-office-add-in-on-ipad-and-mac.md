---
title: Realizar o sideload de suplementos do Office em um iPad ou Mac para teste
description: Teste o suplemento do Office no iPad e no Mac por Sideload.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 7c5e9542c6e6f9abc96defde389b9543421b8529
ms.sourcegitcommit: 604361e55dee45c7a5d34c2fa6937693c154fc24
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/03/2020
ms.locfileid: "47364050"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>Realizar o sideload de suplementos do Office em um iPad ou Mac para teste

Para ver como seu suplemento será executado no Office no iOS, você pode realizar o sideload do manifesto do seu suplemento em um iPad usando o iTunes ou realizar o sideload do manifesto do suplemento diretamente no Office no Mac. Esta ação não permite definir pontos de interrupção e depurar o código do seu suplemento enquanto ele estiver em execução, mas é possível ver como ele se comporta e verificar se a interface do usuário é utilizável e está sendo processada adequadamente.

## <a name="prerequisites-for-office-on-ios"></a>Pré-requisitos do Office no iOS

- Um computador com Windows ou Mac com [iTunes](https://www.apple.com/itunes/download/) instalado.
  > [!IMPORTANT]
  > Se você estiver executando o macOS Catalina, o [iTunes não estará mais disponível](https://support.apple.com/HT210200) , portanto, você deve seguir as instruções na seção [Sideload um suplemento no Excel ou no Word no iPad usando o MacOS,](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) mais adiante neste artigo.

- Um iPad executando o iOS 8,2 ou posterior com o [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) ou [Word](https://apps.apple.com/app/microsoft-word/id586447913) instalado e um cabo de sincronização.

- O arquivo de manifesto .xml para o suplemento que você deseja testar.

## <a name="prerequisites-for-office-on-mac"></a>Pré-requisitos do Office no Mac

- Um Mac executando OS X v10.10 “Yosemite” ou posterior com [Office no Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) instalado.

- Word no Mac versão 15.18 (160109).

- Excel no Mac versão 15.19 (160206).

- PowerPoint no Mac versão 15.24 (160614)

- O arquivo de manifesto .xml para o suplemento que você deseja testar.

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>Sideload um suplemento no Excel ou no Word no iPad usando iTunes

1. Use um cabo de sincronização para conectar seu iPad ao computador. Se você estiver conectando o iPad ao computador pela primeira vez, você será solicitado a **confiar neste computador?**. Escolha **Confiar** para continuar.

2. No iTunes, escolha o ícone do **iPad** abaixo da barra de menus.

3. Em **configurações** no lado esquerdo do iTunes, escolha **aplicativos**.

4. No lado direito do iTunes, role para baixo até **compartilhamento de arquivos**e, em seguida, escolha **Excel** ou **Word** na coluna **suplementos** .

5. Na parte inferior da coluna documentos do **Excel** ou do **Word** , escolha **Adicionar arquivo**e, em seguida, selecione o arquivo manifest. XML do suplemento que você deseja Sideload.

6. Abra o aplicativo Excel ou Word em seu iPad. Se o aplicativo Excel ou Word já estiver em execução, escolha o botão **página inicial** e, em seguida, feche e reinicie o aplicativo.

7. Abra um documento.

8. Escolha **suplementos** na guia **Inserir** . (na guia **Inserir** , talvez seja necessário rolar horizontalmente até que você veja o botão **suplementos** .) O suplemento do suplementos foi feito está disponível para inserção sob o título do **desenvolvedor** na interface do usuário de **suplementos** .

    ![Inserir Suplementos no aplicativo do Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>Sideload um suplemento no Excel ou no Word no iPad com o macOS Catalina

> [!IMPORTANT]
> Com a introdução do macOS Catalina, a [Apple descontinuava o iTunes no Mac](https://support.apple.com/HT210200) e a funcionalidade integrada necessária para os aplicativos do Sideload no **Finder**.

1. Use um cabo de sincronização para conectar seu iPad ao computador. Se você estiver conectando o iPad ao computador pela primeira vez, você será solicitado a **confiar neste computador?**. Escolha **Confiar** para continuar. Você também pode ser perguntado se este é um novo iPad ou se você está restaurando um.

2. No Finder, em **locais**, escolha o ícone do **iPad** abaixo da barra de menus.

3. Na parte superior da janela Localizador, clique em **arquivos**e, em seguida, localize **Excel** ou **Word**.

4. Em uma janela de localizador diferente, arraste e solte o manifest.xml arquivo do suplemento que você deseja carregar no lado do arquivo do **Excel** ou **Word** na primeira janela do Finder.

5. Abra o aplicativo Excel ou Word em seu iPad. Se o aplicativo Excel ou Word já estiver em execução, escolha o botão **página inicial** e, em seguida, feche e reinicie o aplicativo.

6. Abra um documento.

7. Escolha **suplementos** na guia **Inserir** . (na guia **Inserir** , talvez seja necessário rolar horizontalmente até que você veja o botão **suplementos** .) O suplemento do suplementos foi feito está disponível para inserção sob o título do **desenvolvedor** na interface do usuário de **suplementos** .

    ![Inserir Suplementos no aplicativo do Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a>Realizar sideload de um suplemento no Office no Mac

> [!NOTE]
> Para realizar o sideload de um suplemento do Outlook no Mac, confira [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-an-add-in-in-outlook-on-the-desktop).

1. Abra o **terminal** e vá para uma das seguintes pastas onde você salvará o arquivo de manifesto do suplemento. Se a pasta `wef` não existir em seu computador, crie-a.

    - Para o Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Para o Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - Para o PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

2. Abra a pasta no **Finder** usando o comando `open .` (incluindo o ponto ou ponto). Copie o arquivo de manifesto do suplemento nessa pasta.

    ![Pasta Wef no Office no Mac](../images/all-my-files.png)

3. Abra o Word e abra um documento. Reinicie o Word se já estiver em execução.

4. No Word, escolha **Inserir**  >  **suplementos**  >  **meus** suplementos (menu suspenso) e, em seguida, escolha seu suplemento.

    ![Meus Suplementos no Office no Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Aplicativos em que foi feito o sideload não aparecerão na caixa de diálogo Meus Suplementos. Eles só ficam visíveis dentro do menu suspenso (pequena seta para baixo à direita de Meus Suplementos na guia **Inserir**). Os suplementos em que foi feito o sideload são exibidos na lista sob o título **Suplementos do Desenvolvedor** nesse menu.

5. Verifique se o seu suplemento é exibido no Word.

    ![Suplemento do Office exibido no Office no Mac](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>Remover um suplemento do suplementos foi feito

Você pode remover um suplemento suplementos foi feito anteriormente limpando o cache do Office em seu computador. Detalhes sobre como limpar o cache para cada plataforma e aplicativo podem ser encontrados no artigo [limpar o cache do Office](clear-cache.md).

## <a name="see-also"></a>Confira também

- [Depurar suplementos do Office no iPad e no Mac](debug-office-add-ins-on-ipad-and-mac.md)
