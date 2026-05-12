# DataConverter Addin for VBA
In the modern landscape of data exchange, the ability to seamlessly translate external data into a usable VBA format is essential. This Add-in provides a powerful suite of conversion tools—JsonConverter, XMLConverter, and IniConverter—designed to bridge the gap between common data-interchange formats and the Excel VBA environment.

By leveraging the familiar architecture of Excel Collections and Scripting Runtime Dictionaries, this Add-in transforms complex text-based data into intuitive memory structures. This allows developers to interact with external data using standard VBA syntax, making the integration of Web APIs and configuration files both simple and efficient.

## Core Components
1. **JsonConverter:** Modeled after robust Java implementations, this tool is the primary engine for modern Web API integration. It parses JSON strings or physical .json files directly into nested Dictionaries and Collections, allowing for easy key-value pair access.
2. **XMLConverter:** For legacy systems and specialized web services that rely on XML, this converter parses XML text or files into structured memory objects, ensuring compatibility with older Web API standards without the overhead of complex DOM navigation.
3. **IniConverter:** Designed for application settings and local configurations, this tool parses standard .ini files. It provides a lightweight way to manage persistent user settings and environmental variables within your Excel projects.

## Why Use These Converters?
Whether you are fetching real-time data from a modern REST API in JSON format or retrieving structured data from a legacy XML web service, these tools eliminate the need for manual string parsing or "Regex" hacking.

By converting these formats into native Dictionaries and Collections, the Add-in provides a "plug-and-play" experience. You can iterate through data, access specific nodes by name, and manage configuration settings using the standard coding patterns you already know, significantly reducing development time and increasing code reliability.
