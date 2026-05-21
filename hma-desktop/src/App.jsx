import React, { useEffect, useMemo, useState } from "react";
import {
  Activity,
  BarChart3,
  Bot,
  Boxes,
  ChevronDown,
  ChevronRight,
  Clock3,
  Database,
  FileSpreadsheet,
  FolderOpen,
  HelpCircle,
  LayoutGrid,
  Pause,
  Play,
  RefreshCw,
  ShieldCheck,
  Sparkles,
  Wrench
} from "lucide-react";

function GoogleAdsLogo() {
  const configureFrequency = async (freq) => {
    setFrequencyPicker(null);
    await handleAction({
      title: `Configurar informes ${freq}`,
      command: `set-report-frequency:${freq}`
    });
  };

  const runReportNow = async (freq) => {
    setRunReportPicker(null);
    await handleAction({
      title: `Generar informe ${freq} ahora`,
      command: `run-report-now:${freq}`
    });
  };

  return (
    <div className="platform-logo">
      <svg viewBox="0 0 64 64">
        <rect x="12" y="8" width="14" height="42" rx="7" transform="rotate(-28 19 29)" fill="#4285f4" />
        <rect x="29" y="16" width="14" height="38" rx="7" transform="rotate(28 36 35)" fill="#34a853" />
        <circle cx="44" cy="48" r="8" fill="#fbbc05" />
      </svg>
    </div>
  );
}

function MetaLogo() {
  return (
    <div className="platform-logo">
      <svg viewBox="0 0 64 64">
        <path
          d="M10 39c4-17 11-24 19-24 8 0 12 11 16 18 4-8 8-18 16-18 8 0 13 7 13 15 0 8-5 14-13 14-10 0-15-10-20-19-4 8-8 19-17 19-9 0-15-6-15-15 0-3 0-6 1-10Z"
          transform="translate(-9 -3) scale(1.02)"
          fill="none"
          stroke="#27a5ff"
          strokeWidth="6"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
    </div>
  );
}

function LocalLogo() {
  return (
    <div className="platform-logo platform-logo-box">
      <LayoutGrid size={18} />
    </div>
  );
}

function AdminLogo() {
  return (
    <div className="platform-logo platform-logo-box">
      <ShieldCheck size={18} />
    </div>
  );
}

const menuData = [
  {
    key: "google",
    label: "Google Ads",
    logo: GoogleAdsLogo,
    children: [
      {
        key: "google-manual",
        label: "Manual",
        title: "Manual",
        description: "Gestiona conexion y actualizacion manual de Google Ads.",
        actions: [
          { title: "Conectar / editar cuenta", subtitle: "Configura la cuenta Google Ads de un cliente.", icon: Wrench, command: "connect_ads.bat", help: "Abre el flujo para conectar o editar credenciales de Google Ads." },
          { title: "Actualizar metricas manual", subtitle: "Extrae datos manuales para un cliente.", icon: RefreshCw, command: "export_ads.bat", help: "Ejecuta una exportacion manual sin esperar la automatizacion." },
          { title: "Actualizar Google Ads ahora", subtitle: "Actualiza solo Google Ads en todos los clientes.", icon: BarChart3, command: "export_google_all_clients.bat", help: "Corre exportacion global solo de Google Ads." },
          { title: "Ver estado metricas clientes", subtitle: "Revisa datos, CSV y masters por cliente.", icon: Activity, command: "hma_automation_overview.bat", help: "Muestra estado general de clientes y metricas." }
        ]
      },
      {
        key: "google-auto",
        label: "Automatizacion",
        title: "Automatizacion",
        description: "Controla automatizaciones relacionadas con metricas cada 12 horas.",
        actions: [
          { title: "Ver automatizacion de informes", subtitle: "Consulta estado de la tarea automatica.", icon: Clock3, command: "hma_report_frequency_status.bat", help: "Muestra la frecuencia activa, ultima ejecucion y proxima corrida." },
          { title: "Generar informe ahora", subtitle: "Elegí frecuencia y genera el informe manualmente.", icon: RefreshCw, command: "choose-run-report-now", help: "Permite generar un informe manual inmediato con frecuencia 1h, 3h, 5h, 7h, 12h, 1d, 2d o 1w." },
          { title: "Configurar frecuencia de informes", subtitle: "Activa el ciclo completo automatico.", icon: Play, command: "choose-report-frequency", help: "Activa el flujo completo cada 12 horas." },
          { title: "Pausar automatizacion de informes", subtitle: "Pausa la automatizacion completa.", icon: Pause, command: "hma_report_frequency_disable.bat", help: "Pausa la tarea sin borrar datos." },
          { title: "Abrir logs", subtitle: "Abre registros del sistema.", icon: FolderOpen, command: "logs", help: "Abre la carpeta de logs." }
        ]
      }
    ]
  },
  {
    key: "meta",
    label: "Meta Ads",
    logo: MetaLogo,
    children: [
      {
        key: "meta-manual",
        label: "Manual",
        title: "Manual",
        description: "Gestiona conexion y actualizacion manual de Meta Ads.",
        actions: [
          { title: "Conectar / editar cuenta", subtitle: "Configura la cuenta Meta Ads de un cliente.", icon: Wrench, command: "connect_ads.bat", help: "Abre el flujo para conectar o editar credenciales de Meta Ads." },
          { title: "Actualizar metricas manual", subtitle: "Extrae datos manuales para un cliente.", icon: RefreshCw, command: "export_ads.bat", help: "Ejecuta una exportacion manual sin esperar la automatizacion." },
          { title: "Actualizar Meta Ads ahora", subtitle: "Actualiza solo Meta Ads en todos los clientes.", icon: BarChart3, command: "export_meta_all_clients.bat", help: "Corre exportacion global solo de Meta Ads." },
          { title: "Ver estado metricas clientes", subtitle: "Revisa datos, CSV y masters por cliente.", icon: Activity, command: "hma_automation_overview.bat", help: "Muestra estado general de clientes y metricas." }
        ]
      },
      {
        key: "meta-auto",
        label: "Automatizacion",
        title: "Automatizacion",
        description: "Controla automatizaciones relacionadas con metricas cada 12 horas.",
        actions: [
          { title: "Ver automatizacion de informes", subtitle: "Consulta estado de la tarea automatica.", icon: Clock3, command: "hma_report_frequency_status.bat", help: "Muestra la frecuencia activa, ultima ejecucion y proxima corrida." },
          { title: "Generar informe ahora", subtitle: "Elegí frecuencia y genera el informe manualmente.", icon: RefreshCw, command: "choose-run-report-now", help: "Permite generar un informe manual inmediato con frecuencia 1h, 3h, 5h, 7h, 12h, 1d, 2d o 1w." },
          { title: "Configurar frecuencia de informes", subtitle: "Activa el ciclo completo automatico.", icon: Play, command: "choose-report-frequency", help: "Activa el flujo completo cada 12 horas." },
          { title: "Pausar automatizacion de informes", subtitle: "Pausa la automatizacion completa.", icon: Pause, command: "hma_report_frequency_disable.bat", help: "Pausa la tarea sin borrar datos." },
          { title: "Abrir logs", subtitle: "Abre registros del sistema.", icon: FolderOpen, command: "logs", help: "Abre la carpeta de logs." }
        ]
      }
    ]
  },
  {
    key: "local",
    label: "Local",
    logo: LocalLogo,
    children: [
      {
        key: "local-hma",
        label: "HMA actual",
        title: "HMA actual",
        description: "Gestiona el HMA local existente y sus archivos principales.",
        actions: [
          { title: "Ver tareas HMA actual", subtitle: "Consulta tareas y estado local.", icon: Activity, command: "hma_health_check.bat", help: "Ejecuta diagnostico general del sistema." },
          { title: "Abrir HMA_Master.xlsx", subtitle: "Abre el master principal.", icon: FileSpreadsheet, command: "historico\\HMA_Master.xlsx", help: "Abre el Excel maestro principal." }
        ]
      },
      {
        key: "local-masters",
        label: "Masters clientes",
        title: "Masters clientes",
        description: "Construccion y revision de masters por cliente.",
        actions: [
          { title: "Construir masters de todos los clientes", subtitle: "Reconstruye Excels por cliente.", icon: Boxes, command: "build_all_client_masters.bat", help: "Reconstruye HMA_Master.xlsx por cliente usando CSV existentes." },
          { title: "Ver estado metricas clientes", subtitle: "Revisa configuracion y exports.", icon: Activity, command: "hma_automation_overview.bat", help: "Muestra estado de clientes y archivos." },
          { title: "Abrir carpeta clientes", subtitle: "Abre carpeta de clientes.", icon: FolderOpen, command: "clientes", help: "Abre la carpeta donde viven los clientes." }
        ]
      },
      {
        key: "local-backups",
        label: "Backups",
        title: "Backups",
        description: "Proteccion, restauracion y mantenimiento del sistema local.",
        actions: [
          { title: "Crear backup local ahora", subtitle: "Genera ZIP recuperable.", icon: Database, command: "backup_hma_local.bat", help: "Crea un backup local inmediato." },
          { title: "Ver backup semanal", subtitle: "Consulta estado del backup.", icon: Clock3, command: "scripts\\status_weekly_backup_task.ps1", help: "Muestra estado del backup automatico semanal." },
          { title: "Programar backup semanal", subtitle: "Activa backup automatico.", icon: Play, command: "scripts\\setup_weekly_backup_task.ps1", help: "Programa backup semanal." },
          { title: "Pausar backup semanal", subtitle: "Desactiva backup automatico.", icon: Pause, command: "scripts\\disable_weekly_backup_task.ps1", help: "Pausa backup semanal sin borrar backups." },
          { title: "Abrir carpeta backups", subtitle: "Abre carpeta de backups.", icon: FolderOpen, command: "backups", help: "Abre carpeta de ZIPs de backup." },
          { title: "Restaurar backup local", subtitle: "Abre restaurador visual.", icon: RefreshCw, command: "scripts\\restore_hma_backup_gui.ps1", help: "Abre restaurador de backups." }
        ]
      },
      {
        key: "local-logs",
        label: "Logs",
        title: "Logs",
        description: "Acceso a logs, Git y diagnosticos.",
        actions: [
          { title: "Abrir logs", subtitle: "Abre registros del sistema.", icon: FolderOpen, command: "logs", help: "Abre carpeta de logs." },
          { title: "Ver Git status", subtitle: "Consulta cambios pendientes.", icon: Activity, command: "git-status", help: "Muestra estado Git." },
          { title: "Health check sistema", subtitle: "Ejecuta diagnostico general.", icon: ShieldCheck, command: "hma_health_check.bat", help: "Ejecuta health check." }
        ]
      }
    ]
  },
  {
    key: "admin",
    label: "Administrador",
    logo: AdminLogo,
    children: [
      {
        key: "admin-clientes",
        label: "Clientes",
        title: "Clientes",
        description: "Administracion de clientes.",
        actions: [
          { title: "Crear cliente", subtitle: "Crea estructura de cliente.", icon: Play, command: "create_client.bat", help: "Abre flujo para crear cliente." },
          { title: "Ver clientes creados", subtitle: "Lista clientes existentes.", icon: Activity, command: "clients-list", help: "Muestra clientes creados." },
          { title: "Abrir carpeta clientes", subtitle: "Abre carpeta general.", icon: FolderOpen, command: "clientes", help: "Abre carpeta de clientes." }
        ]
      },
      {
        key: "admin-estado",
        label: "Estado / Git",
        title: "Estado / Git",
        description: "Diagnostico, automatizaciones y operaciones globales.",
        actions: [
          { title: "Ver Git status", subtitle: "Consulta cambios pendientes.", icon: Activity, command: "git-status", help: "Muestra estado Git." },
          { title: "Ver todas las automatizaciones", subtitle: "Reporte de tareas programadas.", icon: Boxes, command: "hma_automation_overview.bat", help: "Muestra automatizaciones y estados." },
          { title: "Ejecutar ciclo completo ahora", subtitle: "Exporta, reconstruye y diagnostica.", icon: RefreshCw, command: "hma_run_full_cycle.bat", help: "Ejecuta flujo completo." },
          { title: "Validacion QA completa", subtitle: "Revisa estructura y tareas.", icon: ShieldCheck, command: "hma_qa_validation.bat", help: "Ejecuta validacion QA." }
        ]
      }
    ]
  }
];

function getSelected(data, key) {
  for (const section of data) {
    for (const child of section.children) {
      if (child.key === key) return { section, child };
    }
  }
  return { section: data[0], child: data[0].children[0] };
}

export default function App() {
  const [selectedKey, setSelectedKey] = useState("local-masters");
  const [openSections, setOpenSections] = useState({ google: true, meta: true, local: true, admin: true });
  const [helpItem, setHelpItem] = useState(null);
  const [runResult, setRunResult] = useState(null);
  const [runningTitle, setRunningTitle] = useState("");
  const [systemStatus, setSystemStatus] = useState(null);
  const [frequencyPicker, setFrequencyPicker] = useState(null);
  const [runReportPicker, setRunReportPicker] = useState(null);
  const frequencyOptions = [
    { key: "1h", label: "Cada 1 hora", detail: "Genera informes en Informe_1h." },
    { key: "3h", label: "Cada 3 horas", detail: "Genera informes en Informe_3h." },
    { key: "5h", label: "Cada 5 horas", detail: "Genera informes en Informe_5h." },
    { key: "7h", label: "Cada 7 horas", detail: "Genera informes en Informe_7h." },
    { key: "12h", label: "Cada 12 horas", detail: "Genera informes en Informe_12h." },
    { key: "1d", label: "Cada 1 dia", detail: "Genera informes en Informe_1d." },
    { key: "2d", label: "Cada 2 dias", detail: "Genera informes en Informe_2d." },
    { key: "1w", label: "Cada 1 semana", detail: "Genera informes en Informe_1w." }
  ];

  const loadSystemStatus = async () => {
    try {
      if (window.hma?.getSystemStatus) {
        const status = await window.hma.getSystemStatus();
        setSystemStatus(status);
      }
    } catch (error) {
      setSystemStatus({
        clientsCount: 0,
        backupsCount: 0,
        reportsCount: 0,
        gitClean: false,
        gitStatus: error?.message || String(error),
        activeFrequency: "Error",
        activeTaskName: "No disponible",
        activeTaskState: "Error",
        healthy: false
      });
    }
  };

  useEffect(() => {
    loadSystemStatus();
    const timer = setInterval(loadSystemStatus, 30000);
    return () => clearInterval(timer);
  }, []);

  const { section, child } = useMemo(() => getSelected(menuData, selectedKey), [selectedKey]);

  const selectSection = (sectionItem) => {
    setOpenSections((prev) => ({ ...prev, [sectionItem.key]: true }));
    setSelectedKey(sectionItem.children[0].key);
  };

  const toggleSection = (event, sectionKey) => {
    event.stopPropagation();
    setOpenSections((prev) => ({ ...prev, [sectionKey]: !prev[sectionKey] }));
  };

  const handleAction = async (action) => {
    if (action.command === "choose-report-frequency") {
      setFrequencyPicker(action);
      return;
    }

    if (action.command === "choose-run-report-now") {
      setRunReportPicker(action);
      return;
    }
    setRunningTitle(action.title);
    setRunResult({ title: action.title, command: action.command, ok: null, stdout: "Ejecutando...", stderr: "" });

    try {
      if (!window.hma?.runAction) {
        setRunResult({ title: action.title, command: action.command, ok: false, stdout: "", stderr: "window.hma.runAction no esta disponible. Revisar preload.js." });
        return;
      }

      const result = await window.hma.runAction({
        title: action.title,
        command: action.command
      });

      setRunResult(result);
    } catch (error) {
      setRunResult({ title: action.title, command: action.command, ok: false, stdout: "", stderr: error?.message || String(error) });
    } finally {
      setRunningTitle("");
      loadSystemStatus();
    }
  };

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="sidebar-header">
          <div className="brand-mark"><Bot size={26} /></div>
          <div className="brand-copy">
            <h1>HMA Manager</h1>
            <p>Local AI Operations</p>
          </div>
        </div>

        <div className="sidebar-scroll">
          <div className="sidebar-section-title">PLATAFORMAS</div>

          <nav className="sidebar-nav">
            {menuData.map((sectionItem) => {
              const Logo = sectionItem.logo;
              const isOpen = openSections[sectionItem.key];
              const active = sectionItem.children.some((item) => item.key === selectedKey);

              return (
                <div className="menu-block" key={sectionItem.key}>
                  <button className={`menu-section ${active ? "is-active" : ""}`} onClick={() => selectSection(sectionItem)}>
                    <span className="menu-section-left"><Logo />{sectionItem.label}</span>
                    <span className="menu-toggle" onClick={(event) => toggleSection(event, sectionItem.key)}>
                      {isOpen ? <ChevronDown size={18} /> : <ChevronRight size={18} />}
                    </span>
                  </button>

                  {isOpen && (
                    <div className="submenu">
                      {sectionItem.children.map((sub) => (
                        <button key={sub.key} className={`submenu-item ${selectedKey === sub.key ? "is-selected" : ""}`} onClick={() => setSelectedKey(sub.key)}>
                          <span className="submenu-dot" />
                          <span>{sub.label}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </nav>
        </div>

        <div className="sidebar-footer">
          <div className="mini-status-card">
            <span className="mini-status-dot" />
            <div>
              <strong>Estado del sistema</strong>
              <p>Todo funcionando correctamente</p>
            </div>
          </div>
        </div>
      </aside>

      <main className="main-panel">
        <div className="content-scroll">
          <header className="content-header">
            <div>
              <div className="breadcrumb">{section.label} <span>&gt;</span> {child.label}</div>
              <h2>{child.title}</h2>
              <p>{child.description}</p>
            </div>

            <div className="status-card">
              <div className="status-icon"><ShieldCheck size={22} /></div>
              <div>
                <small>{systemStatus?.gitClean ? "Git limpio" : "Git con cambios"}</small>
                <strong>{systemStatus?.healthy === false ? "Revisar" : "Operativo"}</strong>
              </div>
            </div>
          </header>

          {systemStatus && (
            <section className="live-status-grid">
              <div className="live-status-card">
                <small>Clientes</small>
                <strong>{systemStatus.clientsCount}</strong>
              </div>

              <div className="live-status-card">
                <small>Backups ZIP</small>
                <strong>{systemStatus.backupsCount}</strong>
              </div>

              <div className="live-status-card">
                <small>Informes</small>
                <strong>{systemStatus.reportsCount}</strong>
              </div>

              <div className="live-status-card live-status-card-wide">
                <small>Automatizacion activa</small>
                <strong>{systemStatus.activeFrequency}</strong>
                <span>{systemStatus.nextRunTime ? `Proxima: ${systemStatus.nextRunTime}` : systemStatus.activeTaskState}</span>
              </div>
            </section>
          )}

          <section className="cards-list">
            {child.actions.map((action) => {
              const Icon = action.icon;
              return (
                <div className={`action-card ${runningTitle === action.title ? "is-running" : ""}`} key={action.title}>
                  <button className="action-card-main" onClick={() => handleAction(action)}>
                    <div className="action-icon"><Icon size={22} /></div>
                    <div className="action-copy">
                      <h3>{action.title}</h3>
                      <p>{action.subtitle}</p>
                    </div>
                  </button>

                  <button className="help-button" onClick={() => setHelpItem(action)} title="Ver ayuda">
                    <HelpCircle size={18} />
                  </button>
                </div>
              );
            })}
          </section>
        </div>

        <div className="neural-bg"><span /><span /><span /><span /></div>
      </main>



      {runReportPicker && (
        <div className="modal-backdrop" onClick={() => setRunReportPicker(null)}>
          <div className="help-modal frequency-modal" onClick={(event) => event.stopPropagation()}>
            <div className="help-modal-header">
              <div className="help-modal-icon"><RefreshCw size={20} /></div>
              <div>
                <h4>Generar informe ahora</h4>
                <p>Elegí la frecuencia del informe manual que querés generar.</p>
              </div>
            </div>

            <div className="frequency-grid">
              {frequencyOptions.map((option) => (
                <button key={option.key} className="frequency-option" onClick={() => runReportNow(option.key)}>
                  <strong>{option.label}</strong>
                  <span>{option.detail}</span>
                </button>
              ))}
            </div>

            <div className="help-modal-footer">
              <button className="primary-btn" onClick={() => setRunReportPicker(null)}>Cancelar</button>
            </div>
          </div>
        </div>
      )}

      {frequencyPicker && (
        <div className="modal-backdrop" onClick={() => setFrequencyPicker(null)}>
          <div className="help-modal frequency-modal" onClick={(event) => event.stopPropagation()}>
            <div className="help-modal-header">
              <div className="help-modal-icon"><Clock3 size={20} /></div>
              <div>
                <h4>Configurar frecuencia de informes</h4>
                <p>Elegí cada cuánto querés generar reportes. No hay opción personalizada.</p>
              </div>
            </div>

            <div className="frequency-grid">
              {frequencyOptions.map((option) => (
                <button key={option.key} className="frequency-option" onClick={() => configureFrequency(option.key)}>
                  <strong>{option.label}</strong>
                  <span>{option.detail}</span>
                </button>
              ))}
            </div>

            <div className="help-modal-footer">
              <button className="primary-btn" onClick={() => setFrequencyPicker(null)}>Cancelar</button>
            </div>
          </div>
        </div>
      )}

      {helpItem && (
        <div className="modal-backdrop" onClick={() => setHelpItem(null)}>
          <div className="help-modal" onClick={(event) => event.stopPropagation()}>
            <div className="help-modal-header">
              <div className="help-modal-icon"><HelpCircle size={20} /></div>
              <div>
                <h4>{helpItem.title}</h4>
                <p>Guia rapida</p>
              </div>
            </div>
            <div className="help-modal-body"><p>{helpItem.help}</p></div>
            <div className="help-modal-footer">
              <button className="primary-btn" onClick={() => setHelpItem(null)}>Cerrar</button>
            </div>
          </div>
        </div>
      )}

      {runResult && (
        <div className="modal-backdrop" onClick={() => setRunResult(null)}>
          <div className="help-modal result-modal" onClick={(event) => event.stopPropagation()}>
            <div className="help-modal-header">
              <div className={`help-modal-icon ${runResult.ok === true ? "is-ok" : runResult.ok === false ? "is-error" : ""}`}>
                <Activity size={20} />
              </div>
              <div>
                <h4>{runResult.title}</h4>
                <p>{runResult.command}</p>
              </div>
            </div>

            <div className="result-status">
              {runResult.ok === null ? "EJECUTANDO" : runResult.ok ? "OK" : "ERROR"}
            </div>

            <div className="output-box">
              <strong>Salida</strong>
              <pre>{runResult.stdout || "Sin salida."}</pre>
            </div>

            {runResult.stderr && (
              <div className="output-box output-box-error">
                <strong>Errores / advertencias</strong>
                <pre>{runResult.stderr}</pre>
              </div>
            )}

            <div className="help-modal-footer">
              <button className="primary-btn" onClick={() => setRunResult(null)}>Cerrar</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
