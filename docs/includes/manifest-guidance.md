> [!TIP]
> Se você estiver testando seu suplemento em vários ambientes (por exemplo, em desenvolvimento, preparação, demonstração, etc.), recomendamos que você mantenha um arquivo de manifesto XML diferente para cada ambiente. Em cada arquivo de manifesto, você pode:
> - Especifique as URLs que correspondem ao ambiente.
> - Personalizar os valores de metadados como `DisplayName` e os rótulos dentro `Resources` para indicar o ambiente, para que os usuários finais possam identificar o ambiente correspondente do suplemento do suplementos foi feito. 
> - Personalizar as funções personalizadas `namespace` para indicar o ambiente, se seu suplemento define funções personalizadas.
> 
> Seguindo este guia, você simplificará o processo de teste e evitará problemas que poderiam ocorrer quando um suplemento estiver simultaneamente suplementos foi feito para vários ambientes.