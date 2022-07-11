---
title: Realizar sideload de Suplementos do Office no iPad para teste
description: Teste seu Suplemento do Office no iPad por sideload.
ms.date: 06/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0ba52ae78bed36c4eb8130c714577a1b0899aeb6
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713193"
---
# <a name="sideload-office-add-ins-on-ipad-for-testing"></a>Realizar sideload de Suplementos do Office no iPad para teste

Para ver como o suplemento será executado no Office no iOS, você pode fazer sideload do manifesto do suplemento em um iPad usando o iTunes. Esta ação não permite definir pontos de interrupção e depurar o código do seu suplemento enquanto ele estiver em execução, mas é possível ver como ele se comporta e verificar se a interface do usuário é utilizável e está sendo processada adequadamente.

> [!NOTE]
> Para realizar o sideload de um suplemento do Outlook, confira [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md).

## <a name="prerequisites-for-office-on-ios"></a>Pré-requisitos do Office no iOS

- Um computador com Windows ou Mac com [iTunes](https://www.apple.com/itunes/download/) instalado.
  > [!IMPORTANT]
  > Se você estiver executando o macOS Catalina, o [iTunes](https://support.apple.com/HT210200) não estará mais disponível, portanto, siga as instruções na seção Sideload de um suplemento no [Excel ou word no iPad usando o macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) mais adiante neste artigo.

- Um iPad que executa o iOS 8.2 ou posterior com [o Excel](https://apps.apple.com/app/microsoft-excel/id586683407) ou [Word](https://apps.apple.com/app/microsoft-word/id586447913) instalado e um cabo de sincronização.

- O arquivo de manifesto .xml para o suplemento que você deseja testar.

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>Realizar sideload de um suplemento no Excel ou no Word no iPad usando o iTunes

1. Use um cabo de sincronização para conectar seu iPad ao computador. Se você estiver conectando o iPad ao seu computador pela primeira vez, você será solicitado a confiar **neste computador?** Escolha **Confiar** para continuar.

2. No iTunes, escolha o **ícone do iPad** abaixo da barra de menus.

3. Em **Configurações** no lado esquerdo do iTunes, escolha **Aplicativos**.

4. No lado direito do iTunes, role para baixo até Compartilhamento de Arquivos **e escolha** **Excel** ou **Word** na **coluna Suplementos** .

5. Na parte inferior da coluna **Documentos do Excel** ou **do Word** , escolha Adicionar Arquivo **e, em** seguida, selecione o arquivo .xml manifesto do suplemento que você deseja fazer sideload.

6. Abra o aplicativo Excel ou Word em seu iPad. Se o aplicativo Excel ou Word já estiver em execução, escolha o **botão** Página Inicial e feche e reinicie o aplicativo.

7. Abra um documento.

8. Escolha **Suplementos na** guia Inserir.  (Na guia Inserir, talvez  seja necessário rolar horizontalmente até ver o botão **Suplementos**.) Seu suplemento de sideload está disponível para ser inserido sob o título Desenvolvedor  na interface do usuário **de Suplementos**.

    ![Inserir Suplementos no aplicativo do Excel.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>Fazer sideload de um suplemento no Excel ou no Word no iPad usando o macOS Catalina

> [!IMPORTANT]
> Com a introdução do macOS Catalina, a Apple descontinuou o [iTunes no Mac](https://support.apple.com/HT210200) e a funcionalidade integrada necessária para sideload de aplicativos **no Finder**.

1. Use um cabo de sincronização para conectar seu iPad ao computador. Se você estiver conectando o iPad ao seu computador pela primeira vez, você será solicitado a confiar **neste computador?** Escolha **Confiar** para continuar. Você também pode ser perguntado se este é um novo iPad ou se você está restaurando um.

2. No Localizador, em **Locais**, escolha o **ícone do iPad** abaixo da barra de menus.

3. Na parte superior da janela Localizador, clique em **Arquivos** e localize **o Excel** ou **o Word**.

4. Em outra janela do Finder, arraste e solte o arquivo manifest.xml do suplemento que você deseja carregar lateralmente no arquivo do **Excel** ou **do Word** na primeira janela do Localizador.

5. Abra o aplicativo Excel ou Word em seu iPad. Se o aplicativo Excel ou Word já estiver em execução, escolha o **botão** Página Inicial e feche e reinicie o aplicativo.

6. Abra um documento.

7. Escolha **Suplementos na** guia Inserir.  (Na guia Inserir, talvez  seja necessário rolar horizontalmente até ver o botão **Suplementos**.) Seu suplemento de sideload está disponível para ser inserido sob o título Desenvolvedor  na interface do usuário **de Suplementos**.

    ![Inserir Suplementos no aplicativo do Excel.](../images/excel-insert-add-in.png)

## <a name="remove-a-sideloaded-add-in"></a>Remover um suplemento de sideload

Você pode remover um suplemento com sideload anteriormente limpando o cache do Office em seu computador. Detalhes sobre como limpar o cache para cada plataforma e aplicativo podem ser encontrados no artigo [Limpar o cache do Office](clear-cache.md).

## <a name="see-also"></a>Confira também

- [Realizar sideload de Suplementos do Office no Mac para teste](sideload-an-office-add-in-on-mac.md)
- [Depurar Suplementos do Office em um Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md)
