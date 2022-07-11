---
title: Exigências especiais de suplementos no iPad
description: Conheça alguns requisitos para criar um Suplemento do Office executado em um iPad.
ms.date: 09/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: 17df8855a987bd44e657f6ddfdec9925a979449a
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712990"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>Exigências especiais de suplementos no iPad

Se o suplemento usa apenas APIs do Office com suporte no iPad, os clientes podem instalá-lo em iPads. (Consulte [Especificar aplicativos do Office e requisitos de API](specify-office-hosts-and-api-requirements.md) para obter mais informações.) Se o suplemento for comercializado por meio do *[AppSource](https://appsource.microsoft.com)*, há algumas práticas que você deve seguir para suplementos que podem ser instalados em iPads, além das práticas recomendadas que se aplicam a todos os [Suplementos do Office](../concepts/add-in-development-best-practices.md).

A tabela a seguir lista as tarefas a serem executadas.

> [!NOTE]
> Para obter informações sobre como criar suplementos do Outlook que parecem bons e funcionam bem no Outlook Mobile, consulte [Suplementos para Outlook Mobile](../outlook/outlook-mobile-addins.md).

|Tarefa|Descrição|Recursos|
|:-----|:-----|:-----|
|Atualize seu suplemento para dar suporte ao Office.js versão 1.1.|Atualize os arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no projeto do seu Suplemento do Office para a versão 1.1.|[Atualizar a versão da API e do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Aplicar práticas recomendadas de design do iOS.|Integre perfeitamente a interface do usuário do seu suplemento à experiência para iOS.| Veja a observação abaixo. |
|Otimize seu suplemento para toque.|Torne sua interface do usuário responsiva a entradas de toque, além de mouse e teclado.|[Aplicar os princípios de design da UX](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Torne seu suplemento gratuito.|O Office no iPad é um canal pelo qual você pode atingir mais usuários e promover seus serviços. Esses novos usuários têm potencial para se tornarem seus clientes.|[Política de certificação 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Torne seu comércio de suplementos gratuito no iPad.|Quando ele está em execução no iPad, seu suplemento deve estar livre de compras no aplicativo, ofertas de avaliação, interface do usuário que visam venda adicional para uma versão não gratuita ou links para qualquer loja online em que os usuários possam comprar ou adquirir outros conteúdos, aplicativos ou suplementos. Suas páginas de Política de Privacidade e Termos de Uso também devem estar livres de qualquer interface do usuário de comércio ou links do AppSource.|[Política de certificação 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>Seu suplemento ainda pode ter comércio em outras plataformas. Para fazer isso, teste a propriedade [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member) e suprime todo o comércio quando ela retorna `false`.|
|Envie seu suplemento para o AppSource.|No Partner Center, na página  Configuração do produto, marque a caixa de seleção Tornar meu produto disponível no **iOS e android (** se aplicável) e forneça sua ID de desenvolvedor da Apple nas configurações da conta. Examine o [Contrato de Provedor de](https://go.microsoft.com/fwlink/?linkid=715691) Aplicativos para garantir que você entenda os termos.|[Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> Seu suplemento pode fornecer uma interface do usuário alternativa com base no dispositivo no qual ele está sendo executado. Para detectar se o suplemento está em execução em um iPad, você pode usar as APIs a seguir.
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#office-office-context-touchenabled-member)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member)
>
> Em um iPad, retorna `touchEnabled` `true` e `commerceAllowed` retorna `false`.
>
> Para obter informações sobre as melhores práticas de design de interface do usuário para iPad, consulte [Design para iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/).

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>Práticas recomendadas para o desenvolvimento de Suplementos do Office que podem ser executados no iPad

Aplique as práticas recomendadas a seguir para desenvolver suplementos executados no iPad.

-  **Desenvolva e depure o suplemento no Windows ou Mac e o sideload para um iPad.**

    Você não pode desenvolver o suplemento diretamente em um iPad, mas pode depurá-lo em um computador Windows ou Mac e fazer sideload dele em um iPad para teste. Como um suplemento executado no Office no iOS ou Mac dá suporte às mesmas APIs que um suplemento em execução no Office no Windows, o código do suplemento deve ser executado da mesma maneira nessas plataformas. Para obter detalhes, [consulte Testar e depurar suplementos do Office](../testing/test-debug-office-add-ins.md) e [realizar sideload de suplementos do Office no iPad para teste](../testing/sideload-an-office-add-in-on-ipad.md).

-  **Especifique os requisitos da API no manifesto do seu suplemento ou com verificações da execução.**

    Quando você especificar requisitos de API no manifesto do suplemento, o Office determinará se o aplicativo cliente do Office dá suporte a esses membros da API. Se os membros da API estão disponíveis no aplicativo, seu suplemento estará disponível. Como alternativa, você pode executar uma verificação de runtime para determinar se um método está disponível no aplicativo antes de usá-lo em seu suplemento. Verificações de runtime garantem que o suplemento esteja sempre disponível no aplicativo e fornecem funcionalidade adicional se os métodos estão disponíveis. Para obter mais informações, consulte [Especificar aplicativos do Office e requisitos de API](specify-office-hosts-and-api-requirements.md).
