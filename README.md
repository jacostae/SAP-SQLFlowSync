
# Integración de Datos SAP-SQL con Dashboard en Looker Studio

#### Descripción del Proyecto:
Este proyecto automatiza el proceso de integración de datos entre **_SAP_** y **_Siclo_** (una base de datos SQL), utilizando **scripts en Python** para descargar información desde una transacción específica en SAP, cruzar esta información con datos de despacho, y actualizar una hoja de cálculo en **Google Sheets**. Posteriormente, esta hoja de cálculo alimenta un dashboard en **Looker Studio**, que se actualiza automáticamente para proporcionar visualizaciones en tiempo real.

#### Objetivo del Proyecto:
El objetivo principal de este proyecto es **_optimizar_** y **_automatizar_** el flujo de trabajo que involucra la **extracción de datos** desde SAP, su integración con datos almacenados en Siclo, y la actualización de un **dashboard** en Looker Studio. Esto permite a los usuarios acceder a informes actualizados y visualizaciones de datos en tiempo real sin intervención manual, mejorando la eficiencia y la precisión del proceso.

## Paso a paso para su ejecución


### 1. Configuración proyecto en Google Cloud Platform

Esto se realiza con el fin de obtener las credenciales necesarias para acceder a la hoja de cálculo, en la que se actualizará la información obtenida por el script.

En el siguiente [vídeo](https://www.youtube.com/watch?v=Mz9JG9CUXXY) se observa como se obtiene el archivo:
- credentials.json

Tener en cuenta que el "_robot_" creado en este paso debe tener acceso al archivo o archivos de Google Sheets a actualizar.

![tempsnip](https://github.com/user-attachments/assets/4ca16b2e-2040-45b1-8915-c07b89a88453)

### 2. Revisión usuarios y contraseñas en scripts

En el archivo _scripts.py_ se deben modificar los siguientes campos:

- **sap_user** (Usuario de SAP que tenga acceso a la transacción _Y_CSD_80000073_).
- **sap_password** (Contraseña respectiva).
- **server**, **database**, **username** y **password** para poder acceder a Siclo.

Esto con el fin de poder acceder a SAP y Siclo y obtener la información de cada uno.


### 3. Creación spreadsheet

La _spreadsheet_ debe contar con el nombre que tiene en el código así como las hojas para que el código no genere error al no encontrar dónde guardar los datos.

![hoja_de_calculo](https://github.com/user-attachments/assets/9f1cce4b-9aba-42ff-9123-e36572e43651)

### 4. Creación dashboard en Looker Studio

Teniendo la spreadsheet con los datos podemos crear el dashboard para mejorar la visualización de la información. 

En el siguiente link se puede observar como se realiza la conexión [Conexión hoja de cálculo - Looker Studio](https://support.google.com/looker-studio/answer/6370353?hl=es-419#zippy=%2Csecciones-de-este-art%C3%ADculo)

A continuación un ejemplo de lo que se puede realizar con los datos en Looker Studio:

<img src="https://github.com/user-attachments/assets/a310923e-ffa8-4375-a14b-8e4a9106c87e" alt="Tablero_1" width="700" height="600"/>

<img src="https://github.com/user-attachments/assets/e10d1c01-6f96-4a18-913f-cdc74759dfa8" alt="Tablero_2" width="700" height="500"/>

Este dashboard se envía periodicamente al terminar turno mediante la opción de **_Programar envío_** de Looker Studio.

### 4. Ejecución del main.py con Programador de tareas

Con el fin de evitar la intervención manual en el proceso de actualización del dashboard y los datos, el código se ejecuta de la siguiente manera en _Acciones_:

![image](https://github.com/user-attachments/assets/6b00c3b5-8e0b-47d3-a2ce-1f715d0408ea)

También se puede utilizar el siguinte código xml, guardarlo e importar la tarea:

<details>
  <summary>Ver código XML</summary>

```xml
<Task xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task" version="1.2">
  <RegistrationInfo>
    <Date>2024-03-01T13:48:05.6705751</Date>
    <Author>CO\jacostae</Author>
    <Description>Código para actualizar información SAP y Siclo</Description>
    <URI>\Código_actualización</URI>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <Repetition>
        <Interval>PT4H</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <StartBoundary>2024-03-27T10:59:00-05:00</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-21-4221797372-3623916711-2686236536-24058</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT1M</Interval>
      <Count>2</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\ProgramData\anaconda3\python.exe</Command>
      <Arguments>main.py</Arguments>
      <WorkingDirectory>C:\Users\jacostae\Desktop\Daily_update</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
```
</details>

### Contact Me
<a href="https://co.linkedin.com/in/juan-carlos-acosta-espitia-837735121/"><img alt="LinkedIn" src="https://img.shields.io/badge/LinkedIn-Juan%20Carlos%20Acosta-blue?style=flat-square&logo=linkedin"></a>
<a href="mailto:jc.acosta.espitia@gmail.com"><img alt="Email" src="https://img.shields.io/badge/Gmail-jc.acosta.espitia@gmail.com-red?style=flat-square&logo=gmail"></a>  
