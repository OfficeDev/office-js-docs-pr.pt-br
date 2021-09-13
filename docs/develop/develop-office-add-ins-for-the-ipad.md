---
title: Exigências especiais de suplementos no iPad
description: Saiba alguns requisitos para criar um Office que é executado em um iPad.
ms.date: 09/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: 8a114c5fc4a17ee3f7282321d82ad1faa60d9d71
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148609"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>Exigências especiais de suplementos no iPad

Se o seu add-in usa apenas Office APIs com suporte no iPad, os clientes poderão instalá-lo em iPads. (Consulte [Especificar Office aplicativos e requisitos de API](specify-office-hosts-and-api-requirements.md) para obter mais informações.) Se o add-in for comercializado por meio do *[AppSource](https://appsource.microsoft.com)*, existem algumas práticas que você deve seguir para os complementos que podem ser instalados em iPads, além das práticas recomendadas que se aplicam a todos os [Office Add-ins](../concepts/add-in-development-best-practices.md).

A tabela a seguir lista as tarefas a executar.

> [!NOTE]
> Para obter informações sobre como projetar Outlook de Outlook que parecem bons e funcionam bem no Outlook Mobile, consulte [Add-ins for Outlook Mobile](../outlook/outlook-mobile-addins.md).

|Tarefa|Descrição|Recursos|
|:-----|:-----|:-----|
|Atualize seu suplemento para dar suporte ao Office.js versão 1.1.|Atualize os arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no projeto do seu Suplemento do Office para a versão 1.1.|[Atualizar a versão da API e do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Aplicar práticas recomendadas de design do iOS.|Integre perfeitamente a interface do usuário do seu suplemento à experiência para iOS.| Consulte a observação abaixo. |
|Otimize seu suplemento para toque.|Torne sua interface do usuário responsiva a entradas de toque, além de mouse e teclado.|[Aplicar os princípios de design da UX](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Torne seu suplemento gratuito.|O Office no iPad é um canal pelo qual você pode atingir mais usuários e promover seus serviços. Esses novos usuários têm potencial para se tornarem seus clientes.|[Política de certificação 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Tornar seu comércio de complementos gratuito no iPad.|Quando ele estiver sendo executado no iPad, o seu complemento deve estar livre de compras no aplicativo, ofertas de avaliação, interface do usuário que visa fazer upsell para uma versão não gratuita ou links para qualquer loja online onde os usuários possam comprar ou adquirir outros conteúdos, aplicativos ou complementos. Suas páginas de Política de Privacidade e Termos de Uso também devem estar livres de qualquer interface do usuário do comércio ou links do AppSource.|[Política de certificação 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>Seu complemento ainda pode ter comércio em outras plataformas. Para fazer isso, teste a [propriedade Office.context.commerceAllowed](/javascript/api/office/office.context#commerceAllowed) e suprime todo o comércio quando retornar `false` .|
|Envie seu complemento para AppSource.|No Partner Center,  na página Configuração do produto, marque a caixa de seleção Tornar meu produto disponível no iOS e **Android (se aplicável)** e forneça a ID do desenvolvedor apple nas configurações da conta. Revise o [Contrato de Provedor de Aplicativos](https://go.microsoft.com/fwlink/?linkid=715691) para garantir que você entenda os termos.|[Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> Seu complemento pode atender a uma interface do usuário alternativa com base no dispositivo em que ele está sendo executado. Para detectar se o seu add-in está sendo executado em um iPad, você pode usar as APIs a seguir.
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchEnabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceAllowed)
>
> Em um iPad, `touchEnabled` retorna `true` e retorna `commerceAllowed` `false` .
>
> Para obter informações sobre as melhores práticas de design da interface do usuário para iPad, consulte [Designing for iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/).

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>Práticas recomendadas para Office de Office que podem ser executados em iPad

Aplique as seguintes práticas recomendadas para o desenvolvimento de complementos executados em iPad.

-  **Desenvolva e depure o Windows ou Mac e o coloque em um iPad.**

    Você não pode desenvolver o complemento diretamente em um iPad, mas pode depurá-lo em um computador Windows ou Mac e fazer sideload dele em um iPad para teste. Como um complemento que é executado no Office no iOS ou mac dá suporte às mesmas APIs que um complemento em execução no Office no Windows, o código do seu complemento deve ser executado da mesma maneira nessas plataformas. Para obter detalhes, consulte [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md) and [Sideload Office Add-ins on iPad and Mac for testing](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).

-  **Especifique os requisitos da API no manifesto do seu suplemento ou com verificações da execução.**

    Ao especificar os requisitos de API no manifesto do seu Office, o Office determinará se o aplicativo cliente Office oferece suporte a esses membros da API. Se os membros da API estão disponíveis no aplicativo, seu complemento estará disponível. Como alternativa, você pode executar uma verificação de tempo de execução para determinar se um método está disponível no aplicativo antes de usá-lo no seu complemento. Verificações de tempo de execução garantem que o seu complemento está sempre disponível no aplicativo e fornece funcionalidade adicional se os métodos estão disponíveis. Para obter mais informações, consulte [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).
