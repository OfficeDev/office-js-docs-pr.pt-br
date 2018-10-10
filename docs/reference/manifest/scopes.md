# <a name="scopes-element"></a>Elemento Scopes

Contém as permissões para o Microsoft Graph de que o suplemento precisa. Este elemento é usado pela Office Store para criar uma caixa de diálogo de consentimento. Quando os usuários instalam o suplemento a partir da Office Store, eles são solicitados a conceder ao suplemento as permissões especificas para os dados do Microsoft Graph do usuário.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Tipo  |  Descrição  |
|:-----|:-----|:-----|
|  **Scope**                |  sequência de caracteres     |   O nome de uma permissão para o Microsoft Graph; por exemplo, Files.Read.All. |

## <a name="example"></a>Exemplo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
