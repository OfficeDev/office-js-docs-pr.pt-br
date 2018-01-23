
# <a name="develop-office-add-ins-for-the-ipad"></a>Desenvolver suplementos do Office para iPad


A tabela a seguir lista as tarefas a realizar para desenvolver um Suplemento do Office que será executado no Office para iPad.


|**Tarefa**|**Descrição**|**Recursos**|
|:-----|:-----|:-----|
|Atualize seu suplemento para dar suporte ao Office.js versão 1.1.|Atualize os arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no projeto do seu Suplemento do Office para a versão 1.1.|[O que mudou na API JavaScript para Office](https://dev.office.com/reference/add-ins/what's-changed-in-the-javascript-api-for-office)|
|Aplique as práticas recomendadas de design de interface do usuário.|Integre perfeitamente a interface do usuário do seu suplemento à experiência para iOS.|[Projetar para o iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Aplique as práticas recomendadas de design de suplemento.|Verifique se o suplemento fornece um valor claro, é dedicado e tem um desempenho consistente.|[Práticas recomendadas para desenvolvimento de suplementos do Office](../overview/add-in-development-best-practices.md)|
|Otimize seu suplemento para toque.|Torne sua interface do usuário responsiva a entradas de toque, além de mouse e teclado.|[Aplicar os princípios de design da UX](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#Anchor_3)|
|Torne seu suplemento gratuito.|O Office no iPad é um canal pelo qual você pode atingir mais usuários e promover seus serviços. Esses novos usuários têm potencial para se tornarem seus clientes.|[Política de validação 10.8](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Torne a comercialização do seu suplemento gratuita.|Seu suplemento não deve oferecer compras no aplicativo, ofertas de avaliação, interfaces de usuários com o objetivo de maximizar as vendas nem links para lojas online onde os usuários possam comprar ou adquirir outros conteúdos, aplicativos ou suplementos. Suas páginas de Política de Privacidade e Termos de Uso também não devem ter nenhuma interface de usuário destinada ao comércio ou links para lojas.|[Política de validação 3.4](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Reenvie seu suplemento à Office Store.|No Painel do Vendedor, selecione a caixa **Disponibilizar este suplemento no Catálogo de Suplementos do Office no iPad** e forneça sua ID de desenvolvedor da Apple na caixa ID da Apple. Examine o [Contrato do Provedor de Aplicativo da Office Store](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.htm) para ter certeza de que você o compreendeu.|[Enviar Suplementos do SharePoint e do Office e aplicativos Web do Office 365 à Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)|

Seu suplemento pode permanecer como está para aplicativos do Office que estão sendo executados em outras plataformas. Você também pode fornecer uma interface de usuário diferente com base no navegador/dispositivo em que seu suplemento está sendo executado. Para detectar se seu suplemento está sendo executado em um iPad, você pode usar as seguintes APIs:<ul><li>var isTouchEnabled = [Office.context.touchEnabled](http://dev.office.com/reference/add-ins/shared/office.context.touchenabled)</li><li>var allowCommerce = [Office.context.commerceAllowed](http://dev.office.com/reference/add-ins/shared/office.context.commerceallowed)</li></ul>
    

## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>Práticas recomendadas para desenvolver Suplementos do Office para iOS e Mac

Aplique as seguintes práticas recomendadas para desenvolver suplementos para execução no iOS:


-  **Use o Visual Studio para desenvolver seu suplemento.**
    
    Se você desenvolver seu suplemento com o Visual Studio, é possível [definir pontos de interrupção e depurar seu código](../get-started/create-and-debug-office-add-ins-in-visual-studio.md#Test) em um aplicativo host do Office em execução no Windows antes de realizar o sideload no iPad ou no Mac. Como um suplemento executado no Office para iOS ou no Office para Mac é compatível com as mesmas APIs que um suplemento executado no Office para Windows, o código de seu suplemento deve ser executado da mesma maneira em ambas as plataformas.
    
-  **Especifique os requisitos da API no manifesto do seu suplemento ou com verificações da execução.**
    
    Ao especificar os requisitos da API no manifesto do suplemento, o Office determinará se o aplicativo host é compatível com esses membros da API. Se os membros da API estiverem disponíveis no host, o suplemento ficará disponível nesse aplicativo host. Como alternativa, é possível realizar uma verificação de tempo de execução para determinar se um método está disponível no host antes de usá-lo em seu suplemento. As verificações de tempo de execução garantem que o suplemento sempre esteja disponível no host e proporciona recursos adicionais se os métodos estiverem disponíveis. Para saber mais, consulte [Especificar requisitos de hosts e API para o Office](../overview/specify-office-hosts-and-api-requirements.md).
    
Para ter acesso às práticas recomendadas gerais de desenvolvimento de suplementos, confira [Práticas recomendadas para desenvolvimento de Suplementos do Office](../overview/add-in-development-best-practices.md).


## <a name="additional-resources"></a>Recursos adicionais
<a name="bk_addresources"> </a>


- [Realizar o sideload de um Suplemento do Office no iPad e no Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [Depurar Suplementos do Office no iPad e no Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
