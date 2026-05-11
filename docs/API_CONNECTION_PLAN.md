\# API Connection Plan — HMA



\## 1. Objetivo



Este documento define qué se necesita para pasar el sistema HMA desde modo demo con datos simulados hacia conexión real con plataformas publicitarias.



El sistema actualmente funciona con `DATA\_SOURCE=simulated`.



La próxima fase consiste en conectar:



\- Google Ads API

\- Meta Marketing API



\---



\## 2. Estado actual



| Componente | Estado |

|---|---|

| Demo local | OK |

| GitHub Actions | OK |

| Reporte Markdown | OK |

| CSV exportable | OK |

| JSON exportable | OK |

| Artifact descargable | OK |

| Google Ads API | Pendiente |

| Meta Marketing API | Pendiente |

| Persistencia externa | Pendiente |

| Alertas externas | Pendiente |



\---



\## 3. Principio de seguridad



Las credenciales reales no deben escribirse dentro del código ni subirse al repositorio.



Deben configurarse como:



\- GitHub Actions Secrets;

\- variables seguras del entorno de ejecución;

\- o gestor de secretos de la infraestructura usada.



Nunca deben aparecer en:



\- README;

\- commits;

\- archivos `.py`;

\- archivos `.env` subidos;

\- capturas compartidas públicamente.



\---



\## 4. Credenciales necesarias — Google Ads



Para conectar Google Ads API se necesitan, como mínimo:



```env

GOOGLE\_ADS\_CLIENT\_ID=

GOOGLE\_ADS\_CLIENT\_SECRET=

GOOGLE\_ADS\_DEVELOPER\_TOKEN=

GOOGLE\_ADS\_REFRESH\_TOKEN=

GOOGLE\_ADS\_CUSTOMER\_ID=

