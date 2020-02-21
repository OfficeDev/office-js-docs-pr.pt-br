---
title: Criar suplementos do Outlook para formulários de leitura
description: Suplementos de leitura são suplementos do Outlook que são ativados no painel de leitura ou no inspetor de leitura do Outlook.
ms.date: 04/12/2018
localization_priority: Priority
ms.openlocfilehash: a2a5448261fe6fcd150ed0cabda0184d941334e0
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165705"
---
# <a name="create-outlook-add-ins-for-read-forms"></a>Criar suplementos do Outlook para formulários de leitura

Suplementos de leitura são suplementos do Outlook que são ativados no painel de leitura ou no inspetor de leitura do Outlook. Ao contrário dos suplementos de redação (suplementos do Outlook que são ativados quando um usuário está criando uma mensagem ou um compromisso), os suplementos de leitura ficam disponíveis quando os usuários: 

- Visualizam um email, uma solicitação de reunião, uma resposta de reunião ou um cancelamento da reunião.

   > [!NOTE]
   > O Outlook não ativa suplementos no formulário de leitura para determinados tipos de mensagens, como itens que são anexos de outra mensagem, itens na pasta de rascunhos do Outlook ou itens que estão criptografados ou protegidos de outras maneiras.
    
- Exibem um item de reunião em que o usuário é um participante.
    
- Exibem um item de reunião em que o usuário é o organizador (somente versão RTM do Outlook 2013 e do Exchange 2013).
    
   > [!NOTE]
   > Desde a versão Office 2013 SP1, se o usuário estiver exibindo um item de reunião que o usuário tenha organizado, apenas suplementos redigidos poderão realizar a ativação e estar disponíveis. Os suplementos de leitura não estão mais disponíveis nesse cenário.


Em cada um desses cenários de leitura, o Outlook ativa suplementos quando suas condições de ativação são atendidas e os usuários podem escolher e abrir suplementos ativados na barra de suplemento no Painel de Leitura ou inspetor de leitura. A figura a seguir mostra o suplemento **Bing Mapas** ativado e aberto quando o usuário está lendo uma mensagem que contém um endereço geográfico.


**Painel do suplemento mostrando o suplemento Bing Mapas funcionando, no caso de uma mensagem selecionada do Outlook que contém um endereço**

![Bing Map mail app in Outlook](../images/bing-maps-add-in.jpg)


## <a name="types-of-add-ins-available-in-read-mode"></a>Tipos de suplementos disponíveis no modo de leitura

Suplementos de leitura podem ser uma combinação dos tipos a seguir.

- [Comandos de suplemento para o Outlook](add-in-commands-for-outlook.md)   
- [Suplementos contextuais do Outlook](contextual-outlook-add-ins.md)
    

## <a name="api-features-available-to-read-add-ins"></a>Recursos de API disponíveis para suplementos de leitura

- Para ativar suplementos em formulários de leitura: confira a Tabela 1 em [Especificar regras de ativação em um manifesto](activation-rules.md#specify-activation-rules-in-a-manifest).    
- [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)    
- [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md)    
- [Extrair cadeias de caracteres de entidade de um item do Outlook](extract-entity-strings-from-an-item.md)   
- [Obter anexos de um item do Outlook a partir do servidor](get-attachments-of-an-outlook-item.md)
    

## <a name="see-also"></a>Confira também

- [Escreva seu primeiro suplemento do Outlook](../quickstarts/outlook-quickstart.md)
