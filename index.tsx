import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import { GoogleGenAI } from "@google/genai";
import { 
  ChevronRight, 
  ChevronLeft, 
  CheckCircle, 
  Box, 
  Zap,
  Layout,
  FileOutput,
  MonitorPlay,
  PencilRuler,
  Star,
  ChevronDown,
  Globe,
  Building2,
  Briefcase,
  ShoppingCart,
  FileText,
  Users
} from "lucide-react";

// --- GOOGLE APPS SCRIPT FIX ---
/*
  🔴 IMPORTANTE: Copia TODO este código en tu Google Apps Script para solucionar el error "doGet".
  Recuerda: Dale a "Implementar" > "Nueva versión" después de pegar.

  function doPost(e) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    var row = [
      new Date(),                       // A: Fecha
      data.companyName || "",           // B: Empresa
      data.jobTitle || "",              // C: Cargo
      data.usageType || "",             // D: TIPO DE USO
      
      // Satisfacción (E - K)
      data.satisfactionOverall || 0,    // E
      data.recommendation || 0,         // F
      data.platformStability || 0,      // G
      data.commercialAttention || 0,    // H
      data.technicalSupport || 0,       // I
      data.renewalProcess || 0,         // J
      data.priceValue || 0,             // K
      
      // Estrategia (L - O)
      data.businessImportance || 0,     // L
      data.currentUtilization || 0,     // M
      data.utilizationDesire || 0,      // N
      data.expansionIdeas || "",        // O
      
      // Roadmap (P - R)
      (data.urgentImprovement || []).join(", "), // P
      data.missingFeatures || "",       // Q
      data.featureCriticality || 0,     // R
      
      // --- DATOS DINÁMICOS (S - Z) ---
      
      // E-commerce (S, T)
      data.monthlyOrders || "",         // S
      data.averageOrderValue || "",     // T
      
      // Formulario (U, V, W)
      data.monthlyQuotes || "",         // U
      data.briefingQuality || 0,        // V
      data.clarityImpact || 0,          // W
      
      // Interno (X, Y, Z)
      data.monthlyConfigs || "",        // X
      data.salesConversionImpact || "", // Y
      data.isMeasuring || "",           // Z
      
      // Final (AA)
      data.integrationIdeas || ""       // AA: Comentarios
    ];
    
    sheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({"result":"success"})).setMimeType(ContentService.MimeType.JSON);
  }

  // 👇 ESTO ARREGLA EL ERROR AL ABRIR EL LINK EN EL NAVEGADOR 👇
  function doGet(e) {
    return ContentService.createTextOutput("¡Conexión correcta! El script está activo y esperando datos.");
  }
*/

// --- CONFIGURATION ---

// 🔴 IMPORTANTE: Reemplaza esta URL con la que obtuviste al desplegar tu Google Apps Script
const GOOGLE_SCRIPT_URL: string = "https://script.google.com/macros/s/AKfycbwui479o5ndyujEMf25bAOwuc0QsIOa4maFQAClGispTybDc9SZYobxKGlS-sRNG_7j/exec"; 

// --- Types ---

type UsageType = 'ecommerce' | 'internal' | 'form' | '';

interface SurveyData {
  // Step 0: Identification
  companyName: string;
  jobTitle: string;
  products: string[];
  
  // Step 1: Satisfaction
  satisfactionOverall: number;
  platformStability: number;
  commercialAttention: number;
  technicalSupport: number;
  renewalProcess: number; 
  priceValue: number;
  competitiveness: string; 
  recommendation: number; // NPS

  // Step 2: Usage & Strategy
  businessImportance: number;
  currentUtilization: number;
  utilizationDesire: number;
  expansionIdeas: string;
  integrationIdeas: string;

  // Step 3: Roadmap
  missingFeatures: string;
  featureCriticality: number;
  urgentImprovement: string[];

  // Step 4: Business Context (Dynamic)
  usageType: UsageType; // New segmentation

  // -- Option A: Ecommerce
  monthlyOrders: string;
  averageOrderValue: string;

  // -- Option B: Form / Quotes
  monthlyQuotes: string;
  briefingQuality: number;
  clarityImpact: number;

  // -- Option C: Internal Tool
  monthlyConfigs: string;
  salesConversionImpact: string;
  isMeasuring: string;
}

const initialData: SurveyData = {
  companyName: "",
  jobTitle: "",
  products: ["InstaBox 3D"],
  satisfactionOverall: 3, 
  platformStability: 0,
  commercialAttention: 0,
  technicalSupport: 0,
  renewalProcess: 0,
  priceValue: 0,
  competitiveness: "",
  recommendation: 5,
  businessImportance: 3,
  currentUtilization: 3,
  utilizationDesire: 0,
  expansionIdeas: "",
  integrationIdeas: "",
  missingFeatures: "",
  featureCriticality: 0,
  urgentImprovement: [],
  
  usageType: "",
  // Ecommerce
  monthlyOrders: "",
  averageOrderValue: "",
  // Form
  monthlyQuotes: "",
  briefingQuality: 0,
  clarityImpact: 0,
  // Internal
  monthlyConfigs: "",
  salesConversionImpact: "",
  isMeasuring: ""
};

// --- Translations ---

type Language = 'en' | 'es';

const content = {
  en: {
    headerSubtitle: "Customer Survey",
    step: "Step",
    back: "Back",
    next: "Next Step",
    submit: "Submit Feedback",
    loading: "Analyzing feedback with AI...",
    
    // Step 0 (Identification)
    welcome_title: "Welcome",
    welcome_desc: "Please introduce yourself to start the survey.",
    lbl_company: "Company Name",
    ph_company: "e.g., Acme Corp",
    lbl_role: "Your Job Title / Role",
    ph_role: "e.g., Purchasing Manager, CEO...",

    // Step 1
    s0_title: "Your Experience",
    s0_desc: "Help us improve InstaBox 3D for your business.",
    s0_q1: "Overall Service Quality",
    s0_q1_low: "Very Dissatisfied",
    s0_q1_high: "Extremely Satisfied",
    s0_q2: "Platform Stability",
    s0_q3: "Commercial Support",
    s0_q4: "Technical Support",
    s0_q5: "Admin / License Process",
    s0_q5_opts: ["Very Complicated", "Somewhat Difficult", "Neutral / Okay", "Easy", "Seamless / Automated"],
    s0_q6: "Price-to-Value Ratio",
    s0_q6_opts: ["Poor Value", "Below Average", "Fair Price", "Good Value", "Excellent Investment"],
    s0_nps: "How likely are you to recommend us?",
    s0_nps_low: "Not likely",
    s0_nps_high: "Very likely",

    // Step 2
    s1_title: "Strategic Impact",
    s1_desc: "How does InstaBox 3D fit into your daily workflow?",
    s1_q1: "How critical is InstaBox 3D for your sales?",
    s1_q1_low: "Not Critical",
    s1_q1_high: "Core Business Essential",
    s1_q2: "Current Utilization Level",
    s1_q2_low: "Barely Used",
    s1_q2_high: "Maxed Out",
    s1_q3: "Desire to use more?",
    s1_q3_low: "No Interest",
    s1_q3_high: "Very High",
    s1_q4: "Do you have ideas to extend the use of the configurator? (Optional)",
    s1_q4_ph: "e.g., Use it for internal quoting...",

    // Step 3
    s2_title: "Product Roadmap",
    s2_desc: "Help us prioritize updates for InstaBox 3D.",
    s2_q1: "Which area needs the most urgent improvement?",
    s2_items: {
      editor: "2D Graphic Editor",
      view3d: "3D Visualization Quality",
      anim: "Animations / Interaction",
      ui: "User Interface / Buttons",
      out: "Output Files (.pdf, .dxf)"
    },
    s2_q2: "What specific functionality are you missing?",
    s2_q2_ph: "Describe the feature...",
    s2_q3: "How critical is this feature?",
    s2_q3_low: "Nice to have",
    s2_q3_high: "Absolute Dealbreaker",

    // Step 4 (Business Context - Dynamic)
    s3_title: "Business Insights",
    s3_desc: "Help us understand how you use the tool.",
    s3_usage_q: "How do you mainly use InstaBox 3D?",
    s3_type_internal: "Internal Tool (Sales Team)",
    s3_type_ecommerce: "Ecommerce Integration",
    s3_type_form: "Online Quote Request Form",

    // -- Ecommerce
    s3_ecom_q1: "Average Monthly Orders (Online)",
    s3_ecom_q1_opts: ["0 - 50 orders/mo", "51 - 200 orders/mo", "201 - 1,000 orders/mo", "1,000+ orders/mo"],
    s3_ecom_q2: "Average Order Value (Online)",
    s3_ecom_q2_opts: ["Less than €50", "€50 - €200", "€200 - €500", "More than €500"],

    // -- Form
    s3_form_q1: "Monthly Quote Requests via Web",
    s3_form_q1_opts: ["0 - 20 requests", "21 - 100 requests", "100+ requests"],
    s3_form_q2: "Quality of information received (Briefing)",
    s3_form_q3: "Does it help avoid misunderstandings?",
    
    // -- Internal
    s3_int_q1: "Average Monthly Configurations",
    s3_int_q1_opts: ["0 - 50 configs", "51 - 200 configs", "200+ configs"],
    s3_int_q2: "Impact on Sales Closure Rate",
    s3_int_q2_opts: ["Sin impacto", "Leve mejora", "Mejora significativa", "No lo sé"],
    s3_int_q3: "¿Lo estáis midiendo?",
    s3_int_q3_opts: ["Sí", "No", "Tenemos previsto hacerlo"],

    s3_final_q: "¿Algún comentario final?",
    s3_final_ph: "Peticiones, comentarios generales...",

    // Step 5
    s4_thankyou: "¡Gracias!",
    s4_note: "Respuesta personalizada generada por Gemini",
    s4_restart: "Empezar Nueva Sesión",
    gemini_default: "¡Gracias por tus comentarios! Los revisaremos en breve."
  },
  es: {
    headerSubtitle: "Encuesta de Satisfacción",
    step: "Paso",
    back: "Atrás",
    next: "Siguiente",
    submit: "Enviar Respuesta",
    loading: "Guardando datos y analizando con IA...",
    
    // Step 0
    welcome_title: "Bienvenido/a",
    welcome_desc: "Por favor, identifícate para comenzar la encuesta.",
    lbl_company: "Nombre de la Empresa",
    ph_company: "Ej: Empresa S.L.",
    lbl_role: "Cargo que ocupas",
    ph_role: "Ej: Director de Compras, Gerente...",

    // Step 1
    s0_title: "Tu Experiencia",
    s0_desc: "Ayúdanos a mejorar InstaBox 3D para tu negocio.",
    s0_q1: "Calidad General del Servicio",
    s0_q1_low: "Muy Insatisfecho",
    s0_q1_high: "Muy Satisfecho",
    s0_q2: "Estabilidad de la Plataforma",
    s0_q3: "Atención Comercial",
    s0_q4: "Soporte Técnico",
    s0_q5: "Proceso Admin / Licencias",
    s0_q5_opts: ["Muy Complicado", "Algo Difícil", "Normal / OK", "Fácil", "Automático / Fluido"],
    s0_q6: "Relación Calidad-Precio",
    s0_q6_opts: ["Mala (Muy caro)", "Por debajo de la media", "Precio Justo", "Buena", "Excelente Inversión"],
    s0_nps: "¿Qué tan probable es que nos recomiendes?",
    s0_nps_low: "Poco probable",
    s0_nps_high: "Muy probable",

    // Step 2
    s1_title: "Impacto Estratégico",
    s1_desc: "¿Cómo encaja InstaBox 3D en tu flujo de trabajo?",
    s1_q1: "¿Cómo de crítico es InstaBox 3D para tus ventas?",
    s1_q1_low: "Nada Crítico",
    s1_q1_high: "Esencial / Core",
    s1_q2: "Nivel de Uso Actual",
    s1_q2_low: "Apenas se usa",
    s1_q2_high: "Uso Diario / Máximo",
    s1_q3: "¿Te gustaría sacarle más provecho?",
    s1_q3_low: "Sin interés",
    s1_q3_high: "Mucho interés",
    s1_q4: "¿Tienes ideas para extender el uso del configurador? (Opcional)",
    s1_q4_ph: "Ej: Usarlo para presupuestos internos...",

    // Step 3
    s2_title: "Hoja de Ruta (Roadmap)",
    s2_desc: "Ayúdanos a priorizar las actualizaciones.",
    s2_q1: "¿Qué área necesita una mejora más urgente?",
    s2_items: {
      editor: "Editor Gráfico 2D",
      view3d: "Calidad Visualización 3D",
      anim: "Animaciones / Interacción",
      ui: "Interfaz / Botones",
      out: "Archivos de Salida (.pdf, .dxf)"
    },
    s2_q2: "¿Echas en falta alguna funcionalidad específica?",
    s2_q2_ph: "Describe la funcionalidad...",
    s2_q3: "¿Cómo de crítica es esta funcionalidad?",
    s2_q3_low: "Deseable",
    s2_q3_high: "Imprescindible",

    // Step 4
    s3_title: "Datos de Negocio",
    s3_desc: "Queremos entender cómo usas la herramienta para adaptarnos a ti.",
    s3_usage_q: "¿Qué uso principal le das a InstaBox 3D?",
    s3_type_internal: "Herramienta Interna (Eq. Ventas)",
    s3_type_ecommerce: "Integrado en Ecommerce (Venta)",
    s3_type_form: "Formulario de Presupuestos Online",

    // -- Ecommerce
    s3_ecom_q1: "Pedidos Medios Mensuales (Online)",
    s3_ecom_q1_opts: ["0 - 50 pedidos/mes", "51 - 200 pedidos/mes", "201 - 1,000 pedidos/mes", "1,000+ pedidos/mes"],
    s3_ecom_q2: "Valor Medio del Pedido (Online)",
    s3_ecom_q2_opts: ["Menos de €50", "€50 - €200", "€200 - €500", "Más de €500"],

    // -- Form
    s3_form_q1: "Solicitudes de presupuesto vía Web (Mes)",
    s3_form_q1_opts: ["0 - 20 solicitudes", "21 - 100 solicitudes", "100+ solicitudes"],
    s3_form_q2: "Calidad de la info recibida (Briefing)",
    s3_form_q3: "¿Ayuda a evitar malentendidos?",

    // -- Internal
    s3_int_q1: "Promedio de configuraciones al mes",
    s3_int_q1_opts: ["0 - 50 configs", "51 - 200 configs", "200+ configs"],
    s3_int_q2: "Impacto en tasa de cierre de ventas",
    s3_int_q2_opts: ["Sin impacto", "Leve mejora", "Mejora significativa", "No lo sé"],
    s3_int_q3: "¿Lo estáis midiendo?",
    s3_int_q3_opts: ["Sí", "No", "Tenemos previsto hacerlo"],

    s3_final_q: "¿Algún comentario final?",
    s3_final_ph: "Peticiones, comentarios generales...",

    // Step 5
    s4_thankyou: "¡Gracias!",
    s4_note: "Respuesta personalizada generada por Gemini",
    s4_restart: "Empezar Nueva Sesión",
    gemini_default: "¡Gracias por tus comentarios! Los revisaremos en breve."
  }
};

// --- Components ---

const RatingScale = ({ 
  value, 
  onChange, 
  max = 5, 
  labels = ["Poor", "Excellent"] 
}: { 
  value: number; 
  onChange: (v: number) => void; 
  max?: number; 
  labels?: string[] 
}) => (
  <div className="flex flex-col space-y-2">
    <div className="flex items-center space-x-3">
      {Array.from({ length: max }).map((_, i) => {
        const rating = i + 1;
        return (
          <button
            key={i}
            onClick={() => onChange(rating)}
            className={`w-12 h-12 rounded-xl transition-all duration-200 font-semibold focus:outline-none focus:ring-2 focus:ring-[#0000ff] focus:ring-offset-2 flex items-center justify-center text-lg shadow-sm border
              ${value === rating 
                ? "bg-[#0000ff] border-[#0000ff] text-white shadow-blue-200 scale-105" 
                : "bg-white border-slate-200 text-slate-500 hover:border-blue-300 hover:text-[#0000ff] hover:bg-slate-50"
              }`}
          >
            {rating}
          </button>
        );
      })}
    </div>
    <div className="flex justify-between text-xs text-slate-400 px-1 uppercase tracking-wide w-full max-w-[280px]">
      <span>{labels[0]}</span>
      <span>{labels[1]}</span>
    </div>
  </div>
);

const SliderInput = ({ 
  value, 
  onChange, 
  min = 1, 
  max = 5, 
  labels = ["Low", "High"] 
}: { 
  value: number; 
  onChange: (v: number) => void; 
  min?: number; 
  max?: number; 
  labels?: string[] 
}) => {
  const percentage = ((value - min) / (max - min)) * 100;

  return (
    <div className="w-full py-2">
      <div className="relative w-full h-12 flex items-center">
        <input
          type="range"
          min={min}
          max={max}
          step={1}
          value={value}
          onChange={(e) => onChange(parseInt(e.target.value))}
          className="w-full absolute z-20 opacity-0 cursor-pointer h-full"
        />
        {/* Custom Track */}
        <div className="w-full h-3 bg-slate-200 rounded-full overflow-hidden relative z-10">
          <div 
            className="h-full bg-[#0000ff] transition-all duration-150 ease-out"
            style={{ width: `${percentage}%` }}
          />
        </div>
        {/* Custom Thumb */}
        <div 
          className="absolute h-8 w-8 bg-white border-2 border-[#0000ff] rounded-full shadow-lg z-10 flex items-center justify-center pointer-events-none transition-all duration-150 ease-out"
          style={{ left: `calc(${percentage}% - 16px)` }}
        >
          <span className="text-xs font-bold text-[#0000ff]">{value}</span>
        </div>
      </div>
      <div className="flex justify-between text-xs font-medium text-slate-500 mt-1 uppercase tracking-wider">
        <span>{labels[0]}</span>
        <span>{labels[1]}</span>
      </div>
    </div>
  );
};

const DropdownSelect = ({ 
  value, 
  onChange, 
  options,
  placeholder = "Select..."
}: { 
  value: number | string; 
  onChange: (v: any) => void; 
  options: { val: number | string; label: string }[];
  placeholder?: string;
}) => (
  <div className="relative w-full max-w-sm">
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      className="w-full appearance-none bg-white border border-slate-300 hover:border-[#0000ff] text-slate-700 py-3 px-4 pr-10 rounded-xl leading-tight focus:outline-none focus:ring-2 focus:ring-[#0000ff] focus:border-transparent cursor-pointer shadow-sm transition-all"
    >
      <option value="" disabled>{placeholder}</option>
      {options.map((opt, idx) => (
        <option key={idx} value={opt.val}>
          {opt.label}
        </option>
      ))}
    </select>
    <div className="pointer-events-none absolute inset-y-0 right-0 flex items-center px-4 text-slate-500">
      <ChevronDown size={16} />
    </div>
  </div>
);

// --- Main App ---

const SurveyApp = () => {
  const [lang, setLang] = useState<Language>('es');
  const [step, setStep] = useState(0);
  const [data, setData] = useState<SurveyData>(initialData);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [geminiResponse, setGeminiResponse] = useState<string | null>(null);

  const t = content[lang] || content['es'];
  // Increased total steps to 6 (0=Id, 1=Sat, 2=Strat, 3=Roadmap, 4=Biz, 5=End)
  const totalSteps = 6; 

  const updateData = (key: keyof SurveyData, value: any) => {
    setData(prev => ({ ...prev, [key]: value }));
  };

  // Helper to handle switching types and clearing old data
  const handleUsageTypeChange = (newType: UsageType) => {
    // We create a fresh copy of the current data
    const newData = { ...data, usageType: newType };
    
    // Clear fields that don't belong to the selected type
    if (newType !== 'ecommerce') {
      newData.monthlyOrders = "";
      newData.averageOrderValue = "";
    }
    if (newType !== 'form') {
      newData.monthlyQuotes = "";
      newData.briefingQuality = 0;
      newData.clarityImpact = 0;
    }
    if (newType !== 'internal') {
      newData.monthlyConfigs = "";
      newData.salesConversionImpact = "";
      newData.isMeasuring = "";
    }
    setData(newData);
  };

  const handleNext = () => {
    // Basic validation to ensure they picked a type in Step 4
    if (step === 4 && data.usageType === "") {
      alert("Please select how you use InstaBox 3D.");
      return;
    }
    if (step < totalSteps - 1) setStep(step + 1);
  };

  const handleBack = () => {
    if (step > 0) setStep(step - 1);
  };

  const toggleLang = () => {
    setLang(prev => prev === 'en' ? 'es' : 'en');
  };

  const handleSubmit = async () => {
    setIsSubmitting(true);
    setStep(step + 1);

    // 1. Send data to Google Sheets
    if (GOOGLE_SCRIPT_URL !== "AQUI_TU_URL_DEL_SCRIPT_DE_GOOGLE") {
      try {
        await fetch(GOOGLE_SCRIPT_URL, {
          method: "POST",
          mode: "no-cors", 
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify(data)
        });
      } catch (sheetError) {
        console.error("Error sending to sheets:", sheetError);
      }
    } else {
      console.warn("Google Script URL not configured.");
    }

    // 2. Generate Gemini Response
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
      const prompt = `
        Act as a Customer Success Manager for InstaBox 3D.
        Language of response: ${lang === 'es' ? 'Spanish' : 'English'}.
        User Name/Role: ${data.jobTitle} at ${data.companyName}.
        Usage Type: ${data.usageType}.
        Analyze feedback JSON: ${JSON.stringify(data)}.
        Write a sincere thank you message (100-150 words) addressing them by role if relevant.
        Mention "InstaBox 3D".
        If low satisfaction or complaints, reassure them.
        If high volume or usage, mention partnership value.
        Synthesize a human response in ${lang === 'es' ? 'Spanish' : 'English'}.
      `;

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: prompt,
      });

      setGeminiResponse(response.text || t.gemini_default);
    } catch (error) {
      console.error("Error generating response", error);
      setGeminiResponse(t.gemini_default);
    } finally {
      setIsSubmitting(false);
    }
  };

  const renderStepContent = () => {
    switch (step) {
      case 0: // NEW: Identification
        return (
          <div className="fade-in pb-4">
             <div className="text-center mb-10">
               <h1 className="text-3xl font-bold text-slate-900">{t.welcome_title}</h1>
               <p className="text-slate-500 mt-2 text-lg">{t.welcome_desc}</p>
            </div>
            
            <div className="space-y-8 max-w-md mx-auto">
              <div>
                <label className="block text-base font-semibold text-slate-700 mb-3 flex items-center">
                  <Building2 size={18} className="mr-2 text-[#0000ff]" /> {t.lbl_company}
                </label>
                <input 
                  type="text"
                  className="w-full p-4 border border-slate-300 rounded-xl focus:ring-2 focus:ring-[#0000ff] focus:outline-none text-base shadow-sm"
                  placeholder={t.ph_company}
                  value={data.companyName}
                  onChange={(e) => updateData("companyName", e.target.value)}
                />
              </div>

              <div>
                <label className="block text-base font-semibold text-slate-700 mb-3 flex items-center">
                  <Briefcase size={18} className="mr-2 text-[#0000ff]" /> {t.lbl_role}
                </label>
                <input 
                  type="text"
                  className="w-full p-4 border border-slate-300 rounded-xl focus:ring-2 focus:ring-[#0000ff] focus:outline-none text-base shadow-sm"
                  placeholder={t.ph_role}
                  value={data.jobTitle}
                  onChange={(e) => updateData("jobTitle", e.target.value)}
                />
              </div>
            </div>
          </div>
        );

      case 1: // Satisfaction
        return (
          <div className="fade-in pb-4">
            <div className="text-center mb-10">
               <h1 className="text-3xl font-bold text-slate-900">{t.s0_title}</h1>
               <p className="text-slate-500 mt-2 text-lg">{t.s0_desc}</p>
            </div>
            
            <div className="space-y-10">
              <div className="bg-blue-50/50 p-6 rounded-2xl border border-blue-100">
                <label className="block text-lg font-bold text-slate-800 mb-4">{t.s0_q1}</label>
                <SliderInput 
                  value={data.satisfactionOverall} 
                  onChange={(v) => updateData("satisfactionOverall", v)}
                  min={1}
                  max={10}
                  labels={[t.s0_q1_low, t.s0_q1_high]}
                />
              </div>

              <div className="p-6 rounded-2xl border border-slate-100 bg-white shadow-sm">
                <label className="block text-lg font-bold text-slate-800 mb-4">{t.s0_nps}</label>
                <SliderInput 
                  value={data.recommendation} 
                  onChange={(v) => updateData("recommendation", v)}
                  min={0}
                  max={10}
                  labels={[t.s0_nps_low, t.s0_nps_high]}
                />
              </div>

              <div>
                <label className="block text-base font-semibold text-slate-700 mb-3">{t.s0_q2}</label>
                <RatingScale value={data.platformStability} onChange={(v) => updateData("platformStability", v)} />
              </div>

              <div>
                <label className="block text-base font-semibold text-slate-700 mb-3">{t.s0_q3}</label>
                <RatingScale value={data.commercialAttention} onChange={(v) => updateData("commercialAttention", v)} />
              </div>

              <div>
                <label className="block text-base font-semibold text-slate-700 mb-3">{t.s0_q4}</label>
                <RatingScale value={data.technicalSupport} onChange={(v) => updateData("technicalSupport", v)} />
              </div>

              <div>
                 <label className="block text-base font-semibold text-slate-700 mb-3">{t.s0_q5}</label>
                 <DropdownSelect 
                   value={data.renewalProcess} 
                   onChange={(v) => updateData("renewalProcess", parseInt(v))}
                   options={t.s0_q5_opts.map((label, i) => ({ val: i + 1, label }))}
                   placeholder="Select..."
                 />
              </div>

              <div className="pt-6 border-t border-slate-100">
                <label className="block text-base font-semibold text-slate-700 mb-3">{t.s0_q6}</label>
                <DropdownSelect 
                   value={data.priceValue} 
                   onChange={(v) => updateData("priceValue", parseInt(v))}
                   options={t.s0_q6_opts.map((label, i) => ({ val: i + 1, label }))}
                   placeholder="Select..."
                 />
              </div>
            </div>
          </div>
        );

      case 2: // Strategic Value
        return (
          <div className="fade-in">
            <div className="text-center mb-10">
              <h2 className="text-2xl font-bold text-slate-900">{t.s1_title}</h2>
              <p className="text-slate-500 mt-2">{t.s1_desc}</p>
            </div>

            <div className="space-y-10">
              <div>
                <label className="block text-lg font-bold text-slate-900 mb-4">
                  {t.s1_q1}
                </label>
                <SliderInput 
                  value={data.businessImportance} 
                  onChange={(v) => updateData("businessImportance", v)}
                  labels={[t.s1_q1_low, t.s1_q1_high]}
                />
              </div>

              <div>
                <label className="block text-lg font-bold text-slate-900 mb-4">
                  {t.s1_q2}
                </label>
                <SliderInput 
                  value={data.currentUtilization} 
                  onChange={(v) => updateData("currentUtilization", v)}
                  labels={[t.s1_q2_low, t.s1_q2_high]}
                />
              </div>

              <div>
                <label className="block text-base font-semibold text-slate-900 mb-3">
                  {t.s1_q3}
                </label>
                <RatingScale 
                  value={data.utilizationDesire} 
                  onChange={(v) => updateData("utilizationDesire", v)}
                  labels={[t.s1_q3_low, t.s1_q3_high]}
                />
              </div>
            </div>

            <div className="mt-10 space-y-4 pt-8 border-t border-slate-100">
              <div>
                <label className="block text-base font-medium text-slate-700 mb-2">
                  {t.s1_q4}
                </label>
                <textarea 
                  className="w-full p-4 border border-slate-200 rounded-xl focus:ring-2 focus:ring-[#0000ff] focus:outline-none text-base shadow-sm transition-shadow"
                  rows={2}
                  placeholder={t.s1_q4_ph}
                  value={data.expansionIdeas}
                  onChange={(e) => updateData("expansionIdeas", e.target.value)}
                />
              </div>
            </div>
          </div>
        );

      case 3: // Roadmap
        return (
          <div className="fade-in space-y-8">
            <div className="text-center mb-6">
              <h2 className="text-2xl font-bold text-slate-900">{t.s2_title}</h2>
              <p className="text-slate-500 mt-2">{t.s2_desc}</p>
            </div>

            <div className="space-y-6">
               <label className="block text-base font-bold text-slate-700">{t.s2_q1}</label>
               <div className="grid grid-cols-1 gap-4">
                 {[
                   { id: "2D Editor", icon: PencilRuler, label: t.s2_items.editor },
                   { id: "3D View", icon: Box, label: t.s2_items.view3d },
                   { id: "Animations", icon: MonitorPlay, label: t.s2_items.anim },
                   { id: "UI/UX", icon: Layout, label: t.s2_items.ui },
                   { id: "Output", icon: FileOutput, label: t.s2_items.out }
                 ].map((item) => (
                   <div 
                     key={item.id}
                     onClick={() => {
                        const current = data.urgentImprovement;
                        const isSelected = current.includes(item.id);
                        if (isSelected) updateData("urgentImprovement", current.filter(i => i !== item.id));
                        else updateData("urgentImprovement", [...current, item.id]);
                     }}
                     className={`flex items-center p-4 rounded-xl border-2 cursor-pointer transition-all ${
                       data.urgentImprovement.includes(item.id) 
                         ? "bg-blue-50 border-[#0000ff] text-[#0000ff] shadow-md" 
                         : "bg-white border-slate-100 text-slate-600 hover:border-blue-200 hover:bg-slate-50"
                     }`}
                   >
                     <div className={`p-2 rounded-lg mr-4 ${data.urgentImprovement.includes(item.id) ? "bg-[#0000ff] text-white" : "bg-slate-200 text-slate-500"}`}>
                        <item.icon size={20} />
                     </div>
                     <span className="text-base font-semibold">{item.label}</span>
                     {data.urgentImprovement.includes(item.id) && <CheckCircle size={20} className="ml-auto text-[#0000ff]" />}
                   </div>
                 ))}
               </div>
            </div>

            <div className="pt-8 border-t border-slate-100">
              <label className="block text-base font-medium text-slate-700 mb-3">
                {t.s2_q2}
              </label>
              <textarea 
                className="w-full p-4 border border-slate-200 rounded-xl focus:ring-2 focus:ring-[#0000ff] focus:outline-none text-base shadow-sm"
                rows={3}
                placeholder={t.s2_q2_ph}
                value={data.missingFeatures}
                onChange={(e) => updateData("missingFeatures", e.target.value)}
              />
              {data.missingFeatures && (
                <div className="mt-6 fade-in bg-blue-50 p-6 rounded-2xl border border-blue-100">
                  <label className="block text-sm font-bold text-[#0000ff] mb-4 uppercase tracking-wide">
                    {t.s2_q3}
                  </label>
                  <SliderInput 
                    value={data.featureCriticality} 
                    onChange={(v) => updateData("featureCriticality", v)}
                    labels={[t.s2_q3_low, t.s2_q3_high]}
                  />
                </div>
              )}
            </div>
          </div>
        );

      case 4: // Business Context (DYNAMIC)
        return (
          <div className="fade-in space-y-8">
            <div className="text-center mb-6">
              <h2 className="text-2xl font-bold text-slate-900">{t.s3_title}</h2>
              <p className="text-slate-500 mt-2">{t.s3_desc}</p>
            </div>

            {/* SEGMENTATION SELECTOR */}
            <div className="mb-8">
               <label className="block text-lg font-bold text-slate-800 mb-4 text-center">{t.s3_usage_q}</label>
               <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                  {[
                    { id: 'internal', label: t.s3_type_internal, icon: Users },
                    { id: 'ecommerce', label: t.s3_type_ecommerce, icon: ShoppingCart },
                    { id: 'form', label: t.s3_type_form, icon: FileText },
                  ].map((type) => (
                     <div
                        key={type.id}
                        onClick={() => handleUsageTypeChange(type.id as UsageType)}
                        className={`cursor-pointer p-4 rounded-xl border-2 flex flex-col items-center justify-center text-center transition-all h-32 hover:scale-[1.02]
                          ${data.usageType === type.id
                             ? "border-[#0000ff] bg-blue-50 text-[#0000ff] shadow-md"
                             : "border-slate-100 bg-white text-slate-500 hover:border-blue-200"
                          }
                        `}
                     >
                        <type.icon size={28} className="mb-2" />
                        <span className="text-sm font-bold leading-tight">{type.label}</span>
                     </div>
                  ))}
               </div>
            </div>

            {/* CONDITIONAL QUESTIONS */}
            <div className="space-y-8 bg-slate-50 p-6 rounded-2xl border border-slate-100 animate-[fadeIn_0.3s_ease-in-out]">
              
              {/* ECOMMERCE */}
              {data.usageType === 'ecommerce' && (
                <>
                   <div>
                    <label className="block text-base font-semibold text-slate-700 mb-3">{t.s3_ecom_q1}</label>
                    <DropdownSelect 
                      value={data.monthlyOrders} 
                      onChange={(v) => updateData("monthlyOrders", v)}
                      options={t.s3_ecom_q1_opts.map(o => ({ val: o, label: o }))}
                      placeholder="Select..."
                    />
                  </div>
                  <div>
                    <label className="block text-base font-semibold text-slate-700 mb-3">{t.s3_ecom_q2}</label>
                    <DropdownSelect 
                      value={data.averageOrderValue} 
                      onChange={(v) => updateData("averageOrderValue", v)}
                      options={t.s3_ecom_q2_opts.map(o => ({ val: o, label: o }))}
                      placeholder="Select..."
                    />
                  </div>
                </>
              )}

              {/* FORM / QUOTES */}
              {data.usageType === 'form' && (
                <>
                   <div>
                    <label className="block text-base font-semibold text-slate-700 mb-3">{t.s3_form_q1}</label>
                    <DropdownSelect 
                      value={data.monthlyQuotes} 
                      onChange={(v) => updateData("monthlyQuotes", v)}
                      options={t.s3_form_q1_opts.map(o => ({ val: o, label: o }))}
                      placeholder="Select..."
                    />
                  </div>
                  <div>
                    <label className="block text-base font-semibold text-slate-700 mb-3">{t.s3_form_q2}</label>
                    <RatingScale value={data.briefingQuality} onChange={(v) => updateData("briefingQuality", v)} labels={["Poor Info", "Perfect Info"]} />
                  </div>
                  <div>
                    <label className="block text-base font-semibold text-slate-700 mb-3">{t.s3_form_q3}</label>
                    <RatingScale value={data.clarityImpact} onChange={(v) => updateData("clarityImpact", v)} labels={["No help", "Crucial"]} />
                  </div>
                </>
              )}

              {/* INTERNAL TOOL */}
              {data.usageType === 'internal' && (
                <>
                   <div>
                    <label className="block text-base font-semibold text-slate-700 mb-3">{t.s3_int_q1}</label>
                    <DropdownSelect 
                      value={data.monthlyConfigs} 
                      onChange={(v) => updateData("monthlyConfigs", v)}
                      options={t.s3_int_q1_opts.map(o => ({ val: o, label: o }))}
                      placeholder="Select..."
                    />
                  </div>
                  <div>
                    <label className="block text-base font-semibold text-slate-700 mb-3">{t.s3_int_q2}</label>
                    <DropdownSelect 
                      value={data.salesConversionImpact} 
                      onChange={(v) => updateData("salesConversionImpact", v)}
                      options={t.s3_int_q2_opts.map(o => ({ val: o, label: o }))}
                      placeholder="Select..."
                    />
                  </div>
                  <div>
                    <label className="block text-base font-semibold text-slate-700 mb-3">{t.s3_int_q3}</label>
                    <DropdownSelect 
                      value={data.isMeasuring} 
                      onChange={(v) => updateData("isMeasuring", v)}
                      options={t.s3_int_q3_opts.map(o => ({ val: o, label: o }))}
                      placeholder="Select..."
                    />
                  </div>
                </>
              )}

              {data.usageType === "" && (
                <div className="text-center text-slate-400 italic py-4">
                  Select a usage type above to see relevant questions.
                </div>
              )}
            </div>

            <div className="pt-4 border-t border-slate-100">
              <label className="block text-base font-semibold text-slate-700 mb-3">
                {t.s3_final_q}
              </label>
              <textarea 
                className="w-full p-4 border border-slate-200 rounded-xl focus:ring-2 focus:ring-[#0000ff] focus:outline-none text-base shadow-sm"
                rows={2}
                placeholder={t.s3_final_ph}
                value={data.integrationIdeas} 
                onChange={(e) => updateData("integrationIdeas", e.target.value)}
              />
            </div>
          </div>
        );

      case 5: // Conclusion
        return (
          <div className="text-center space-y-6 fade-in flex flex-col items-center justify-center min-h-[300px]">
            {isSubmitting ? (
              <>
                <div className="w-16 h-16 border-4 border-slate-200 border-t-[#0000ff] rounded-full animate-spin"></div>
                <p className="text-slate-600 animate-pulse font-medium text-lg mt-4">{t.loading}</p>
              </>
            ) : (
              <>
                <div className="w-20 h-20 bg-blue-50 text-[#0000ff] rounded-full flex items-center justify-center mb-6 shadow-sm">
                  <Star size={40} fill="currentColor" />
                </div>
                <h2 className="text-3xl font-bold text-slate-900">{t.s4_thankyou}</h2>
                <div className="bg-white border border-slate-100 shadow-xl p-8 rounded-2xl max-w-lg mx-auto text-left relative overflow-hidden">
                  <div className="absolute top-0 left-0 w-2 h-full bg-[#0000ff]"></div>
                  <p className="text-slate-700 leading-relaxed whitespace-pre-wrap font-medium text-lg">
                    {geminiResponse}
                  </p>
                  <div className="mt-6 flex items-center text-xs text-slate-400 border-t border-slate-100 pt-4">
                    <Zap size={14} className="mr-1 text-yellow-500" />
                    <span>{t.s4_note}</span>
                  </div>
                </div>
                <button 
                  onClick={() => window.location.reload()}
                  className="mt-10 px-8 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-sm font-bold transition-colors"
                >
                  {t.s4_restart}
                </button>
              </>
            )}
          </div>
        );

      default:
        return null;
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4 font-sans">
      <div className="w-full max-w-2xl bg-white rounded-3xl shadow-2xl overflow-hidden flex flex-col min-h-[700px] border border-slate-100">
        
        {/* Progress Bar */}
        <div className="h-1.5 bg-slate-100 w-full">
          <div 
            className="h-full bg-[#0000ff] transition-all duration-500 ease-out"
            style={{ width: `${((step + 1) / totalSteps) * 100}%` }}
          />
        </div>

        {/* Header */}
        <div className="p-6 md:p-8 border-b border-slate-100 flex justify-between items-center bg-white/95 backdrop-blur-sm sticky top-0 z-20">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center text-white font-bold text-sm shadow-lg bg-[#0000ff]">
              <Box size={22} />
            </div>
            <div>
              <span className="block font-bold text-slate-800 tracking-tight text-lg leading-tight">InstaBox 3D</span>
              <span className="text-xs text-slate-400 font-medium">{t.headerSubtitle}</span>
            </div>
          </div>
          
          <div className="flex items-center space-x-3">
            <button 
              onClick={toggleLang}
              className="flex items-center space-x-1 px-3 py-1.5 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold transition-colors"
            >
              <Globe size={14} />
              <span>{lang.toUpperCase()}</span>
            </button>

            <div className="text-xs font-bold text-[#0000ff] bg-blue-50 px-3 py-1.5 rounded-full uppercase tracking-wider hidden sm:block">
              {t.step} {step + 1} / {totalSteps}
            </div>
          </div>
        </div>

        {/* Content Area */}
        <div className="flex-1 p-6 md:p-10 overflow-y-auto">
          {renderStepContent()}
        </div>

        {/* Footer Navigation */}
        {step < totalSteps - 1 && (
          <div className="p-6 md:p-8 border-t border-slate-100 bg-slate-50 flex justify-between items-center z-20">
            <button
              onClick={handleBack}
              disabled={step === 0}
              className={`flex items-center px-5 py-3 text-sm font-bold rounded-xl transition-all
                ${step === 0 
                  ? "text-slate-300 cursor-not-allowed" 
                  : "text-slate-600 hover:bg-white hover:shadow-md hover:text-slate-900"
                }`}
            >
              <ChevronLeft size={18} className="mr-1" /> {t.back}
            </button>
            
            <button
              onClick={step === totalSteps - 2 ? handleSubmit : handleNext}
              className="flex items-center px-8 py-3 bg-[#0000ff] hover:bg-blue-700 text-white text-sm font-bold rounded-xl shadow-lg hover:shadow-xl hover:-translate-y-0.5 transition-all focus:ring-4 focus:ring-blue-200"
            >
              {step === totalSteps - 2 ? t.submit : t.next}
              {step !== totalSteps - 2 && <ChevronRight size={18} className="ml-1" />}
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

const container = document.getElementById("root");
if (container) {
  const root = createRoot(container);
  root.render(<SurveyApp />);
}