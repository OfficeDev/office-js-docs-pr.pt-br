---
title: Desenvolver suplementos do Office para iPad
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: df67b772bbac04eb3380bc4d8ecae9acaaaadc0c
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="develop-office-add-ins-for-the-ipad"></a>Desenvolver suplementos do Office para iPad


A tabela a seguir lista as tarefas a realizar para desenvolver um Suplemento do Office que ser? executado no Office para iPad.


|**Tarefa**|**Descri??o**|**Recursos**|
|:-----|:-----|:-----|
|Atualize seu suplemento para dar suporte ao Office.js vers?o 1.1.|Atualize os arquivos de JavaScript (Office.js e arquivos .js espec?ficos do aplicativo) e o arquivo de valida??o de manifesto de suplemento usados no projeto do seu Suplemento do Office para a vers?o 1.1.|[O que mudou na API JavaScript para Office](https://dev.office.com/reference/add-ins/what's-changed-in-the-javascript-api-for-office)|
|Aplique as pr?ticas recomendadas de design de interface do usu?rio.|Integre perfeitamente a interface do usu?rio do seu suplemento ? experi?ncia para iOS.|[Projetar para o iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Aplique as pr?ticas recomendadas de design de suplemento.|Verifique se o suplemento fornece um valor claro, ? dedicado e tem um desempenho consistente.|[Pr?ticas recomendadas para desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)|
|Otimize seu suplemento para toque.|Torne sua interface do usu?rio responsiva a entradas de toque, al?m de mouse e teclado.|[Aplicar os princ?pios de design da UX](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Torne seu suplemento gratuito.|O Office no iPad ? um canal pelo qual voc? pode atingir mais usu?rios e promover seus servi?os. Esses novos usu?rios t?m potencial para se tornarem seus clientes.|[Pol?tica de valida??o 10.8](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#10-apps-and-add-ins-utilize-supported-capabilities)|
|Torne a comercializa??o do seu suplemento gratuita.|Seu suplemento n?o deve oferecer compras no aplicativo, ofertas de avalia??o, interfaces de usu?rios com o objetivo de maximizar as vendas nem links para lojas online onde os usu?rios possam comprar ou adquirir outros conte?dos, aplicativos ou suplementos. Suas p?ginas de Pol?tica de Privacidade e Termos de Uso tamb?m n?o devem ter nenhuma interface de usu?rio destinada ao com?rcio ou links para o AppSource.|[Pol?tica de valida??o 3.4](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#3-apps-and-add-ins-can-sell-additional-features-or-content-through-purchases-within-the-app-or-add-in)|
|Enviar solu??es ao AppSource|No Painel do Vendedor, selecione a caixa **Disponibilizar este suplemento no Cat?logo de Suplementos do Office no iPad** e forne?a sua ID de desenvolvedor da Apple na caixa ID da Apple. Examine o [Contrato do Provedor de Aplicativo do AppSource](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.htm) para ter certeza de que voc? o compreendeu.|[Disponibilizar suas solu??es no AppSource e no Office](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)|

Seu suplemento pode permanecer como est? para aplicativos do Office que est?o sendo executados em outras plataformas. Voc? tamb?m pode fornecer uma interface de usu?rio diferente com base no navegador/dispositivo em que seu suplemento est? sendo executado. Para detectar se seu suplemento est? sendo executado em um iPad, voc? pode usar as seguintes APIs:
- var isTouchEnabled = [Office.context.touchEnabled](https://dev.office.com/reference/add-ins/shared/office.context.touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](https://dev.office.com/reference/add-ins/shared/office.context.commerceallowed)
    

## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>Pr?ticas recomendadas para desenvolver Suplementos do Office para iOS e Mac

Aplique as seguintes pr?ticas recomendadas para desenvolver suplementos para execu??o no iOS:


-  **Use o Visual Studio para desenvolver seu suplemento.**
    
    Se voc? desenvolver seu suplemento com o Visual Studio, ? poss?vel [definir pontos de interrup??o e depurar seu c?digo](../develop/create-and-debug-office-add-ins-in-visual-studio.md) em um aplicativo host do Office em execu??o no Windows antes de realizar o sideload no iPad ou no Mac. Como um suplemento executado no Office para iOS ou no Office para Mac ? compat?vel com as mesmas APIs que um suplemento executado no Office para Windows, o c?digo de seu suplemento deve ser executado da mesma maneira em ambas as plataformas.
    
-  **Especifique os requisitos da API no manifesto do seu suplemento ou com verifica??es da execu??o.**
    
    Ao especificar os requisitos da API no manifesto do suplemento, o Office determinar? se o aplicativo host ? compat?vel com esses membros da API. Se os membros da API estiverem dispon?veis no host, o suplemento ficar? dispon?vel nesse aplicativo host. Como alternativa, ? poss?vel realizar uma verifica??o de tempo de execu??o para determinar se um m?todo est? dispon?vel no host antes de us?-lo em seu suplemento. As verifica??es de tempo de execu??o garantem que o suplemento sempre esteja dispon?vel no host e proporciona recursos adicionais se os m?todos estiverem dispon?veis. Para saber mais, consulte [Especificar requisitos de hosts e API para o Office](specify-office-hosts-and-api-requirements.md).
    
Para ter acesso ?s pr?ticas recomendadas gerais de desenvolvimento de suplementos, confira [Pr?ticas recomendadas para desenvolvimento de Suplementos do Office](../concepts/add-in-development-best-practices.md).


## <a name="see-also"></a>Veja tamb?m

- [Realizar sideload de um suplemento do Office no iPad e no Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Depurar suplementos do Office no iPad e no Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
