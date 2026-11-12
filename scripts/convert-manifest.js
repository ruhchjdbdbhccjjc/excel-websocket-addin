const fs = require('fs');
const path = require('path');

// Read the JSON manifest
const manifestPath = path.join(__dirname, '..', 'appPackage', 'manifest.json');
const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));

// Extract key values
const id = manifest.id;
const version = manifest.version;
const displayName = manifest.name.short;
const description = manifest.description.short;
const providerName = manifest.developer.name;
const supportUrl = manifest.developer.websiteUrl;
const defaultLocale = 'en-US';

const extension = manifest.extensions[0];
const taskPaneRuntime = extension.runtimes.find(r => r.id === 'SharedRuntime' || r.id === 'TaskPaneRuntime');
const sourceLocation = taskPaneRuntime?.code?.page || 'https://localhost:3000/taskpane.html';

const xml = `<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" 
  xsi:type="TaskPaneApp">
  <Id>${id}</Id>
  <Version>${version}</Version>
  <ProviderName>${providerName}</ProviderName>
  <DefaultLocale>${defaultLocale}</DefaultLocale>
  <DisplayName DefaultValue="${displayName}"/>
  <Description DefaultValue="${description}"/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-80.png"/>
  <SupportUrl DefaultValue="${supportUrl}"/>
  <AppDomains>
    <AppDomain>https://contoso.com</AppDomain>
    <AppDomain>https://localhost:3000</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${sourceLocation}"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <Runtimes>
          <Runtime resid="Taskpane.Url" lifetime="long" />
        </Runtimes>
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="Support.Url"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="ExcelWebSocketGroup">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="IconConnected.16x16"/>
                  <bt:Image size="32" resid="IconConnected.32x32"/>
                  <bt:Image size="80" resid="IconConnected.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="btnExcelWebSocket">
                  <Label resid="Taskpane.Label"/>
                  <Supertip>
                    <Title resid="Taskpane.Label"/>
                    <Description resid="Taskpane.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="IconConnected.16x16"/>
                    <bt:Image size="32" resid="IconConnected.32x32"/>
                    <bt:Image size="80" resid="IconConnected.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
        <bt:Image id="IconConnected.16x16" DefaultValue="https://localhost:3000/assets/icon-connected-16.png"/>
        <bt:Image id="IconConnected.32x32" DefaultValue="https://localhost:3000/assets/icon-connected-32.png"/>
        <bt:Image id="IconConnected.80x80" DefaultValue="https://localhost:3000/assets/icon-connected-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="${sourceLocation}"/>
        <bt:Url id="Support.Url" DefaultValue="${supportUrl}"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="Excel WebSocket"/>
        <bt:String id="GetStarted.Title" DefaultValue="Excel WebSocket Loaded"/>
        <bt:String id="Taskpane.Label" DefaultValue="Excel WebSocket"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Excel WebSocket add-in loaded successfully."/>
        <bt:String id="Taskpane.Tooltip" DefaultValue="Open Excel WebSocket panel and view connection status."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>`;

const outputPath = path.join(__dirname, '..', 'manifest.xml');
fs.writeFileSync(outputPath, xml, 'utf8');

console.log('Manifest XML generated successfully at:', outputPath);
