import React, { useState, useRef, useEffect } from 'react';
import { 
  Download, FileText, Upload, Briefcase, GraduationCap, Award, Sparkles, 
  Code, HeartPulse, BookOpen, Hammer, CheckCircle, Settings, Plus, 
  Trash2, ChevronDown, ChevronUp, Loader2, RefreshCw, Sliders, Globe, 
  Linkedin, MapPin, Mail, Phone, Eye, Edit3, ShieldAlert
} from 'lucide-react';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  ShadingType
} from 'docx';
import html2pdf from 'html2pdf.js';
import { Capacitor } from '@capacitor/core';
import { Filesystem, Directory } from '@capacitor/filesystem';
import { Share } from '@capacitor/share';

// Helper to convert Blob to Base64 for Capacitor saving
const blobToBase64 = (blob) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const base64String = reader.result.toString().split(',')[1];
      resolve(base64String);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
};


// Sample Emma Larsen data matching the schema
const emmaLarsenData = {
  name: "EMMA LARSEN",
  title: "Creative Director",
  contact: {
    phone: "(555) 019-2834",
    email: "emma.larsen@email.com",
    location: "Seattle, WA",
    linkedin: "linkedin.com/in/emmalarsen",
    website: "emmalarsen.design"
  },
  summary: {
    tagline: "Multidisciplinary Creative Director with 8+ years of experience leading cross-functional teams to deliver award-winning brand strategies and digital experiences.",
    bullets: [
      "Directed design strategy for high-profile clients, resulting in an average 35% increase in user engagement.",
      "Spearheaded the redesign of core digital platforms, modernizing the user experience for over 10M active users."
    ],
    skills: ["Brand Strategy", "UI/UX Design", "Creative Direction", "Team Leadership", "Adobe Creative Suite", "Figma", "Typography", "Motion Design"]
  },
  experience: [
    {
      company: "Studio Nord",
      location: "Seattle, WA",
      title: "Creative Director",
      start: "Jan 2022",
      end: "Present",
      summary: "Lead a team of 12 designers and writers to create cohesive brand experiences across web, mobile, and print.",
      bullets: [
        "Oversee the creative direction of all client campaigns, ensuring alignment with brand values and business goals.",
        "Mentored junior and mid-level designers, fostering a culture of innovation and collaboration.",
        "Reduced project delivery times by 20% through the implementation of agile design workflows."
      ]
    },
    {
      company: "Pixel Craft",
      location: "Portland, OR",
      title: "Senior UI/UX Designer",
      start: "Mar 2018",
      end: "Dec 2021",
      summary: "Designed user-centered interfaces for web and mobile applications, collaborating closely with product and engineering teams.",
      bullets: [
        "Conducted user research and usability testing to inform design decisions and iterate on product features.",
        "Created wireframes, prototypes, and high-fidelity mockups for e-commerce and SaaS platforms.",
        "Established the company's first design system, improving consistency and design-to-development handoff."
      ]
    }
  ],
  education: [
    {
      school: "University of Washington",
      location: "Seattle, WA",
      degree: "Bachelor of Fine Arts in Interaction Design",
      date: "2017",
      honors: "Cum Laude, Dean's List"
    }
  ],
  certifications: [
    {
      name: "Human-Centered Design Certificate",
      org: "IDEO u",
      date: "2020"
    },
    {
      name: "Certified Scrum Product Owner (CSPO)",
      org: "Scrum Alliance",
      date: "2019"
    }
  ],
  projects: [
    {
      name: "Brand Identity for GreenTech",
      description: "Developed a comprehensive brand system, including logo, guidelines, and website, for a renewable energy startup.",
      bullets: [
        "Created a modern, sustainable brand image that helped secure $5M in Series A funding.",
        "Designed a responsive website with custom illustrations and animations."
      ]
    }
  ],
  awards: [
    "Red Dot Design Award (2023)",
    "Webby Award Winner - Best Visual Design (2021)"
  ]
};

const initialResumeState = {
  name: "",
  title: "",
  contact: {
    phone: "",
    email: "",
    location: "",
    linkedin: "",
    website: ""
  },
  summary: {
    tagline: "",
    bullets: [],
    skills: []
  },
  experience: [],
  education: [],
  certifications: [],
  projects: [],
  awards: []
};

// Client-side regex-based parser
const parseResumeText = (text) => {
  const lines = text.split('\n').map(line => line.trim()).filter(line => line.length > 0);
  const data = {
    name: "",
    title: "",
    contact: { phone: "", email: "", location: "", linkedin: "", website: "" },
    summary: { tagline: "", bullets: [], skills: [] },
    experience: [],
    education: [],
    certifications: [],
    projects: [],
    awards: []
  };

  let currentSection = 'header';
  let currentItem = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const upperLine = line.toUpperCase();

    // Section detection
    if (upperLine.includes("PROFESSIONAL TITLE") || upperLine === "TITLE") {
      currentSection = 'title';
      continue;
    } else if (upperLine.includes("PROFESSIONAL OVERVIEW") || upperLine.includes("SUMMARY") || upperLine === "PROFILE" || upperLine === "OBJECTIVE") {
      currentSection = 'summary';
      continue;
    } else if (upperLine.includes("WORK EXPERIENCE") || upperLine.includes("EXPERIENCE") || upperLine.includes("EMPLOYMENT") || upperLine.includes("HISTORY")) {
      currentSection = 'experience';
      currentItem = null;
      continue;
    } else if (upperLine.includes("EDUCATION") || upperLine.includes("ACADEMIC")) {
      currentSection = 'education';
      currentItem = null;
      continue;
    } else if (upperLine.includes("SKILLS") || upperLine.includes("EXPERTISE") || upperLine.includes("TECHNOLOGIES")) {
      currentSection = 'skills';
      continue;
    } else if (upperLine.includes("CERTIFICATIONS") || upperLine.includes("LICENSES") || upperLine.includes("CREDENTIALS")) {
      currentSection = 'certifications';
      currentItem = null;
      continue;
    } else if (upperLine.includes("PROJECTS")) {
      currentSection = 'projects';
      currentItem = null;
      continue;
    } else if (upperLine.includes("AWARDS") || upperLine.includes("HONORS")) {
      currentSection = 'awards';
      continue;
    }

    // Process line according to active section
    if (currentSection === 'header') {
      if (!data.name) {
        data.name = line;
      } else {
        const emailMatch = line.match(/\S+@\S+\.\S+/);
        const phoneMatch = line.match(/\+?\d[\d -()]{7,}\d/);
        const linkedinMatch = line.toLowerCase().match(/linkedin\.com\/\S+/);

        if (line.includes('|')) {
          const parts = line.split('|').map(p => p.trim());
          parts.forEach(part => {
            if (part.includes('@')) data.contact.email = part;
            else if (part.match(/\d/) && !part.includes('.')) data.contact.phone = part;
            else if (part.toLowerCase().includes('linkedin.com') || part.toLowerCase().includes('linkedin')) data.contact.linkedin = part;
            else if (part.includes('.com') || part.includes('.design') || part.includes('.io') || part.includes('http')) data.contact.website = part;
            else data.contact.location = part;
          });
        } else {
          if (emailMatch) {
            data.contact.email = emailMatch[0];
          } else if (phoneMatch) {
            data.contact.phone = phoneMatch[0];
          } else if (linkedinMatch) {
            data.contact.linkedin = line;
          } else if (line.includes('.com') || line.includes('.design') || line.includes('.io') || line.includes('http')) {
            data.contact.website = line;
          } else {
            // Check if it's location (shorter text) or Professional Title
            if (line.length < 35 && line.includes(',')) {
              data.contact.location = line;
            } else if (!data.title && line.length < 40) {
              data.title = line;
            }
          }
        }
      }
    } else if (currentSection === 'title') {
      data.title = line;
      currentSection = 'header';
    } else if (currentSection === 'summary') {
      if (line.startsWith('-') || line.startsWith('*') || line.startsWith('•')) {
        data.summary.bullets.push(line.replace(/^[-*•]\s*/, ''));
      } else {
        if (data.summary.tagline) {
          data.summary.tagline += " " + line;
        } else {
          data.summary.tagline = line;
        }
      }
    } else if (currentSection === 'skills') {
      const skillsList = line.split(/[,;|•]/).map(s => s.trim()).filter(s => s.length > 0);
      data.summary.skills = [...data.summary.skills, ...skillsList];
    } else if (currentSection === 'experience') {
      if (line.startsWith('-') || line.startsWith('*') || line.startsWith('•')) {
        if (currentItem) {
          currentItem.bullets.push(line.replace(/^[-*•]\s*/, ''));
        }
      } else {
        const dateRegex = /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{4})\s*\d{0,4}\s*[-–—]\s*(Present|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{4})\s*\d{0,4}/i;
        const dateMatch = line.match(dateRegex);

        if (dateMatch || line.includes(' - ') || line.includes(' | ')) {
          if (!currentItem || dateMatch) {
            currentItem = {
              company: "",
              location: "",
              title: "",
              start: "",
              end: "",
              summary: "",
              bullets: []
            };
            data.experience.push(currentItem);
          }

          if (dateMatch) {
            const datePart = dateMatch[0];
            const parts = datePart.split(/[-–—]/).map(p => p.trim());
            currentItem.start = parts[0] || "";
            currentItem.end = parts[1] || "";
            
            let cleaned = line.replace(dateMatch[0], '').replace(/[()]/g, '').trim();
            if (cleaned.includes('|') || cleaned.includes('-')) {
              const compParts = cleaned.split(/[-|]/).map(p => p.trim());
              currentItem.company = compParts[0];
              currentItem.title = compParts[1] || "";
            } else {
              currentItem.title = cleaned;
            }
          } else {
            const parts = line.split(/[-|]/).map(p => p.trim());
            currentItem.company = parts[0] || "";
            if (parts[1]) currentItem.location = parts[1];
          }
        } else {
          if (currentItem) {
            if (!currentItem.title) {
              currentItem.title = line;
            } else if (!currentItem.summary) {
              currentItem.summary = line;
            } else {
              currentItem.summary += " " + line;
            }
          } else {
            currentItem = {
              company: line,
              location: "",
              title: "",
              start: "",
              end: "",
              summary: "",
              bullets: []
            };
            data.experience.push(currentItem);
          }
        }
      }
    } else if (currentSection === 'education') {
      if (line.startsWith('-') || line.startsWith('*') || line.startsWith('•') || line.toLowerCase().startsWith('honors:')) {
        if (currentItem) {
          currentItem.honors = (currentItem.honors ? currentItem.honors + ", " : "") + line.replace(/^[-*•]\s*/, '').replace(/honors:\s*/i, '');
        }
      } else {
        const dateRegex = /\b(19|20)\d{2}\b/;
        const dateMatch = line.match(dateRegex);

        if (!currentItem || dateMatch || line.includes(' - ') || line.includes(' | ')) {
          currentItem = {
            school: "",
            location: "",
            degree: "",
            date: "",
            honors: ""
          };
          data.education.push(currentItem);
        }

        if (dateMatch) {
          currentItem.date = dateMatch[0];
        }

        if (line.includes('|') || line.includes('-')) {
          const parts = line.split(/[-|]/).map(p => p.trim());
          currentItem.school = parts[0] || "";
          if (parts[1]) currentItem.degree = parts[1].replace(currentItem.date, '').replace(/[()]/g, '').trim();
        } else {
          if (!currentItem.school) {
            currentItem.school = line;
          } else if (!currentItem.degree) {
            currentItem.degree = line.replace(currentItem.date, '').trim();
          }
        }
      }
    } else if (currentSection === 'certifications') {
      const cleanedLine = line.replace(/^[-*•]\s*/, '');
      const dateRegex = /\b(19|20)\d{2}\b/;
      const dateMatch = cleanedLine.match(dateRegex);
      let dateStr = "";
      let textStr = cleanedLine;
      if (dateMatch) {
        dateStr = dateMatch[0];
        textStr = cleanedLine.replace(dateMatch[0], '').replace(/[()]/g, '').trim();
      }

      const parts = textStr.split(',');
      data.certifications.push({
        name: parts[0]?.trim() || textStr,
        org: parts[1]?.trim() || "",
        date: dateStr
      });
    } else if (currentSection === 'projects') {
      if (line.startsWith('-') || line.startsWith('*') || line.startsWith('•')) {
        if (currentItem) {
          currentItem.bullets.push(line.replace(/^[-*•]\s*/, ''));
        }
      } else {
        currentItem = {
          name: line,
          description: "",
          bullets: []
        };
        data.projects.push(currentItem);
      }
    } else if (currentSection === 'awards') {
      data.awards.push(line.replace(/^[-*•]\s*/, ''));
    }
  }

  return data;
};

export default function App() {
  const [resumeText, setResumeText] = useState("");
  const [apiKey, setApiKey] = useState("");
  const [resumeData, setResumeData] = useState(emmaLarsenData); // Pre-fill with Emma Larsen
  const [isParsing, setIsParsing] = useState(false);
  const [error, setError] = useState(null);
  
  // Theme Selection
  const [activeCategory, setActiveCategory] = useState("corporate");
  const [activeVariant, setActiveVariant] = useState("boardroom");
  
  // Creative theme specific accent color
  const [creativeAccent, setCreativeAccent] = useState("#A85432"); // Default Clay

  // Accordion Form Open Sections
  const [openSections, setOpenSections] = useState({
    contact: true,
    summary: false,
    experience: false,
    education: false,
    skills: false,
    certifications: false,
    projects: false,
    awards: false
  });

  const previewRef = useRef();

  // Page Break Detection (On-screen indicator)
  const [pageBreaks, setPageBreaks] = useState([]);
  
  const updatePageBreaks = () => {
    if (previewRef.current) {
      const element = previewRef.current;
      const heightPx = element.scrollHeight;
      // In CSS, 11in = 1056px at standard 96 DPI
      const pageHeightPx = 1056; 
      const pageCount = Math.ceil(heightPx / pageHeightPx);
      
      const breaks = [];
      for (let i = 1; i < pageCount; i++) {
        breaks.push(i * pageHeightPx);
      }
      setPageBreaks(breaks);
    }
  };

  useEffect(() => {
    updatePageBreaks();
    // Observe resize of the preview element
    if (typeof ResizeObserver !== 'undefined' && previewRef.current) {
      const observer = new ResizeObserver(() => {
        updatePageBreaks();
      });
      observer.observe(previewRef.current);
      return () => observer.disconnect();
    }
  }, [resumeData, activeCategory, activeVariant, creativeAccent]);

  const toggleSection = (section) => {
    setOpenSections(prev => ({ ...prev, [section]: !prev[section] }));
  };

  const handleTrySample = () => {
    setResumeData(emmaLarsenData);
    setResumeText(
`EMMA LARSEN
emma.larsen@email.com | (555) 019-2834 | Seattle, WA
linkedin.com/in/emmalarsen | emmalarsen.design

PROFESSIONAL TITLE
Creative Director

PROFESSIONAL OVERVIEW
Multidisciplinary Creative Director with 8+ years of experience leading cross-functional teams to deliver award-winning brand strategies and digital experiences.
- Directed design strategy for high-profile clients, resulting in an average 35% increase in user engagement.
- Spearheaded the redesign of core digital platforms, modernizing the user experience for over 10M active users.

WORK EXPERIENCE
Studio Nord - Seattle, WA
Creative Director (Jan 2022 - Present)
Lead a team of 12 designers and writers to create cohesive brand experiences across web, mobile, and print.
- Oversee the creative direction of all client campaigns, ensuring alignment with brand values and business goals.
- Mentored junior and mid-level designers, fostering a culture of innovation and collaboration.
- Reduced project delivery times by 20% through the implementation of agile design workflows.

Pixel Craft - Portland, OR
Senior UI/UX Designer (Mar 2018 - Dec 2021)
Designed user-centered interfaces for web and mobile applications, collaborating closely with product and engineering teams.
- Conducted user research and usability testing to inform design decisions and iterate on product features.
- Created wireframes, prototypes, and high-fidelity mockups for e-commerce and SaaS platforms.
- Established the company's first design system, improving consistency and design-to-development handoff.

EDUCATION
University of Washington - Seattle, WA
Bachelor of Fine Arts in Interaction Design (2017)
Honors: Cum Laude, Dean's List

CERTIFICATIONS
- Human-Centered Design Certificate, IDEO u (2020)
- Certified Scrum Product Owner (CSPO), Scrum Alliance (2019)

PROJECTS
Brand Identity for GreenTech
Developed a comprehensive brand system, including logo, guidelines, and website, for a renewable energy startup.
- Created a modern, sustainable brand image that helped secure $5M in Series A funding.
- Designed a responsive website with custom illustrations and animations.

AWARDS
- Red Dot Design Award (2023)
- Webby Award Winner - Best Visual Design (2021)`
    );
    // Open some accordion sections
    setOpenSections({
      contact: true,
      summary: true,
      experience: true,
      education: true,
      skills: true,
      certifications: true,
      projects: true,
      awards: true
    });
  };

  const handleParse = async () => {
    if (!resumeText.trim()) {
      setError("Please paste your resume text first.");
      return;
    }

    setIsParsing(true);
    setError(null);

    try {
      if (apiKey.trim()) {
        const response = await fetch('https://api.anthropic.com/v1/messages', {
          method: 'POST',
          headers: {
            'x-api-key': apiKey.trim(),
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json',
            'dangerously-allow-browser': 'true'
          },
          body: JSON.stringify({
            model: 'claude-3-5-sonnet-20241022',
            max_tokens: 2000,
            messages: [
              {
                role: 'user',
                content: `You are a resume parser. Extract the following structured JSON from the pasted resume text.
Return ONLY valid JSON, no preamble, no markdown fences.

Schema:
{
  "name": "string",
  "title": "string",
  "contact": {
    "phone": "string",
    "email": "string",
    "location": "string (City, ST)",
    "linkedin": "string (URL or handle)",
    "website": "string"
  },
  "summary": {
    "tagline": "string (one short italic line, e.g. 'Multi-faceted executive with expertise in:')",
    "bullets": ["string"],
    "skills": ["string"]
  },
  "experience": [{
    "title": "string",
    "company": "string",
    "location": "string",
    "start": "string (e.g. 'Jan 2022')",
    "end": "string ('Present' or end date)",
    "summary": "string (optional paragraph)",
    "bullets": ["string"]
  }],
  "education": [{
    "degree": "string",
    "school": "string",
    "location": "string",
    "date": "string",
    "honors": "string (optional)"
  }],
  "certifications": [{
    "name": "string",
    "org": "string",
    "date": "string"
  }],
  "projects": [{ "name": "string", "description": "string", "bullets": ["string"] }],
  "awards": ["string"]
}

If a field is missing, use empty string or empty array. Do not invent content.

Resume text:
${resumeText}`
              }
            ]
          })
        });

        if (!response.ok) {
          throw new Error(`Anthropic API call failed: ${response.statusText}`);
        }

        const data = await response.json();
        const responseText = data.content[0].text;
        const parsed = JSON.parse(responseText);
        setResumeData(parsed);
        
        // Expand accordion sections
        setOpenSections({
          contact: true,
          summary: true,
          experience: true,
          education: true,
          skills: true,
          certifications: true,
          projects: true,
          awards: true
        });
      } else {
        // Fallback local parsing
        await new Promise(resolve => setTimeout(resolve, 1000));
        const parsed = parseResumeText(resumeText);
        setResumeData(parsed);
        
        setOpenSections({
          contact: true,
          summary: true,
          experience: true,
          education: true,
          skills: true,
          certifications: true,
          projects: true,
          awards: true
        });
      }
    } catch (err) {
      console.error(err);
      setError("We couldn't parse this — paste the structured fields manually");
      // Populate with blank schema, reveal edit panel
      setResumeData(initialResumeState);
      setOpenSections({
        contact: true,
        summary: true,
        experience: true,
        education: true,
        skills: true,
        certifications: true,
        projects: true,
        awards: true
      });
    } finally {
      setIsParsing(false);
    }
  };

  // State update utilities
  const updateContact = (field, value) => {
    setResumeData(prev => ({
      ...prev,
      contact: { ...prev.contact, [field]: value }
    }));
  };

  const updateSummary = (field, value) => {
    setResumeData(prev => ({
      ...prev,
      summary: { ...prev.summary, [field]: value }
    }));
  };

  const handleAddExperience = () => {
    setResumeData(prev => ({
      ...prev,
      experience: [
        ...prev.experience,
        { company: "", location: "", title: "", start: "", end: "", summary: "", bullets: ["", "", ""] }
      ]
    }));
  };

  const handleRemoveExperience = (index) => {
    setResumeData(prev => ({
      ...prev,
      experience: prev.experience.filter((_, i) => i !== index)
    }));
  };

  const handleUpdateExperience = (index, field, value) => {
    setResumeData(prev => {
      const newExp = [...prev.experience];
      newExp[index] = { ...newExp[index], [field]: value };
      return { ...prev, experience: newExp };
    });
  };

  const handleUpdateExperienceBullet = (expIndex, bulletIndex, value) => {
    setResumeData(prev => {
      const newExp = [...prev.experience];
      const newBullets = [...newExp[expIndex].bullets];
      newBullets[bulletIndex] = value;
      newExp[expIndex] = { ...newExp[expIndex], bullets: newBullets };
      return { ...prev, experience: newExp };
    });
  };

  const handleAddEducation = () => {
    setResumeData(prev => ({
      ...prev,
      education: [
        ...prev.education,
        { school: "", location: "", degree: "", date: "", honors: "" }
      ]
    }));
  };

  const handleRemoveEducation = (index) => {
    setResumeData(prev => ({
      ...prev,
      education: prev.education.filter((_, i) => i !== index)
    }));
  };

  const handleUpdateEducation = (index, field, value) => {
    setResumeData(prev => {
      const newEdu = [...prev.education];
      newEdu[index] = { ...newEdu[index], [field]: value };
      return { ...prev, education: newEdu };
    });
  };

  const handleAddCertification = () => {
    setResumeData(prev => ({
      ...prev,
      certifications: [
        ...prev.certifications,
        { name: "", org: "", date: "" }
      ]
    }));
  };

  const handleRemoveCertification = (index) => {
    setResumeData(prev => ({
      ...prev,
      certifications: prev.certifications.filter((_, i) => i !== index)
    }));
  };

  const handleUpdateCertification = (index, field, value) => {
    setResumeData(prev => {
      const newCerts = [...prev.certifications];
      newCerts[index] = { ...newCerts[index], [field]: value };
      return { ...prev, certifications: newCerts };
    });
  };

  const handleAddProject = () => {
    setResumeData(prev => ({
      ...prev,
      projects: [
        ...prev.projects,
        { name: "", description: "", bullets: ["", ""] }
      ]
    }));
  };

  const handleRemoveProject = (index) => {
    setResumeData(prev => ({
      ...prev,
      projects: prev.projects.filter((_, i) => i !== index)
    }));
  };

  const handleUpdateProject = (index, field, value) => {
    setResumeData(prev => {
      const newProj = [...prev.projects];
      newProj[index] = { ...newProj[index], [field]: value };
      return { ...prev, projects: newProj };
    });
  };

  const handleUpdateProjectBullet = (projIndex, bulletIndex, value) => {
    setResumeData(prev => {
      const newProj = [...prev.projects];
      const newBullets = [...newProj[projIndex].bullets];
      newBullets[bulletIndex] = value;
      newProj[projIndex] = { ...newProj[projIndex], bullets: newBullets };
      return { ...prev, projects: newProj };
    });
  };

  // Switch template variants when category changes
  const selectCategory = (category) => {
    setActiveCategory(category);
    if (category === "corporate") setActiveVariant("boardroom");
    else if (category === "tech") setActiveVariant("builder");
    else if (category === "creative") setActiveVariant("editorial");
    else if (category === "healthcare") setActiveVariant("practitioner");
    else if (category === "academic") setActiveVariant("scholar");
    else if (category === "trades") setActiveVariant("operator");
  };

  // PDF Export via html2pdf
  const handleDownloadPDF = async () => {
    const element = previewRef.current;
    if (!element) return;

    const filename = `${resumeData.name.replace(/\s+/g, '_')}_resume.pdf`;
    const opt = {
      margin: 0,
      filename: filename,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { format: 'letter', unit: 'in', orientation: 'portrait' }
    };

    if (Capacitor.isNativePlatform()) {
      try {
        const pdfBlob = await html2pdf().set(opt).from(element).outputPdf('blob');
        const base64Data = await blobToBase64(pdfBlob);
        
        const result = await Filesystem.writeFile({
          path: filename,
          data: base64Data,
          directory: Directory.Cache
        });

        await Share.share({
          title: 'Save PDF Resume',
          url: result.uri
        });
      } catch (error) {
        console.error('Error generating or sharing PDF on native:', error);
        alert('Failed to generate or share PDF: ' + error.message);
      }
    } else {
      html2pdf().set(opt).from(element).save();
    }
  };

  // Word Document Export via docx
  const handleDownloadWord = async () => {
    const isSerif = activeCategory === "corporate" || activeCategory === "creative" || activeCategory === "academic";
    const fontName = isSerif ? "EB Garamond" : "Arial"; // Arial matches standard system sans-serif for Word
    
    const children = [];

    // Helper to add paragraph with standard fonts
    const addPara = (text, options = {}) => {
      return new Paragraph({
        alignment: options.alignment || AlignmentType.LEFT,
        spacing: { before: options.before || 0, after: options.after || 60 },
        children: [
          new TextRun({
            text: text,
            font: fontName,
            size: options.size || 20, // 10pt
            bold: options.bold || false,
            italic: options.italic || false,
            color: options.color || "000000",
            characterSpacing: options.characterSpacing || 0
          })
        ]
      });
    };

    // Header: Name
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [
        new TextRun({
          text: resumeData.name.toUpperCase(),
          font: fontName,
          size: 32, // 16pt
          bold: true,
          characterSpacing: 16
        })
      ]
    }));

    // Header: Title
    if (resumeData.title) {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: resumeData.title.toUpperCase(),
            font: fontName,
            size: 20, // 10pt
            color: "666666"
          })
        ]
      }));
    }

    // Contact row (Pipe-separated, Centered)
    const contactParts = [
      resumeData.contact.phone,
      resumeData.contact.email,
      resumeData.contact.location,
      resumeData.contact.linkedin,
      resumeData.contact.website
    ].filter(Boolean);
    
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      border: {
        bottom: {
          color: "D3D3D3",
          space: 6,
          value: BorderStyle.SINGLE,
          size: 6 // 0.75 pt
        }
      },
      children: [
        new TextRun({
          text: contactParts.join("  |  "),
          font: fontName,
          size: 18 // 9pt
        })
      ]
    }));

    // Section Creator
    const addSectionHeader = (title) => {
      children.push(new Paragraph({
        spacing: { before: 240, after: 120 },
        shading: {
          fill: "F3F4F6", // light grey banner background
          type: ShadingType.CLEAR
        },
        children: [
          new TextRun({
            text: `  ${title.toUpperCase()}  `,
            font: fontName,
            size: 22,
            bold: true,
            color: "333333"
          })
        ]
      }));
    };

    // Summary Section
    if (resumeData.summary.tagline || resumeData.summary.bullets?.length > 0) {
      addSectionHeader("Professional Overview");
      
      if (resumeData.summary.tagline) {
        children.push(new Paragraph({
          spacing: { after: 120 },
          children: [
            new TextRun({
              text: resumeData.summary.tagline,
              font: fontName,
              size: 20,
              italic: true
            })
          ]
        }));
      }

      resumeData.summary.bullets.forEach(bullet => {
        if (bullet.trim()) {
          children.push(new Paragraph({
            bullet: { level: 0 },
            spacing: { after: 60 },
            children: [
              new TextRun({
                text: bullet,
                font: fontName,
                size: 20
              })
            ]
          }));
        }
      });
    }

    // Experience Section
    if (resumeData.experience?.length > 0) {
      addSectionHeader("Work Experience");

      resumeData.experience.forEach(entry => {
        // Company | Location + Date (Right aligned tab stop)
        children.push(new Paragraph({
          tabStops: [{ type: "right", position: 10080 }], // right tab stop at 7 inches (10080 dxa)
          spacing: { before: 120, after: 40 },
          children: [
            new TextRun({
              text: `${entry.company} | ${entry.location}`,
              font: fontName,
              size: 20,
              bold: true
            }),
            new TextRun({
              text: `\t${entry.start} – ${entry.end}`,
              font: fontName,
              size: 20,
              bold: true
            })
          ]
        }));

        // Role Title
        if (entry.title) {
          children.push(new Paragraph({
            spacing: { after: 80 },
            children: [
              new TextRun({
                text: entry.title,
                font: fontName,
                size: 20,
                italic: true
              })
            ]
          }));
        }

        // Job Description Paragraph
        if (entry.summary) {
          children.push(new Paragraph({
            spacing: { after: 100 },
            children: [
              new TextRun({
                text: entry.summary,
                font: fontName,
                size: 20
              })
            ]
          }));
        }

        // Job Bullets
        entry.bullets.forEach(bullet => {
          if (bullet.trim()) {
            children.push(new Paragraph({
              bullet: { level: 0 },
              spacing: { after: 60 },
              children: [
                new TextRun({
                  text: bullet,
                  font: fontName,
                  size: 20
                })
              ]
            }));
          }
        });
      });
    }

    // Education Section
    if (resumeData.education?.length > 0) {
      addSectionHeader("Education");

      resumeData.education.forEach(edu => {
        children.push(new Paragraph({
          tabStops: [{ type: "right", position: 10080 }],
          spacing: { before: 120, after: 40 },
          children: [
            new TextRun({
              text: `${edu.school} | ${edu.location}`,
              font: fontName,
              size: 20,
              bold: true
            }),
            new TextRun({
              text: `\t${edu.date}`,
              font: fontName,
              size: 20,
              bold: true
            })
          ]
        }));

        children.push(new Paragraph({
          spacing: { after: 40 },
          children: [
            new TextRun({
              text: edu.degree,
              font: fontName,
              size: 20
            })
          ]
        }));

        if (edu.honors) {
          children.push(new Paragraph({
            spacing: { after: 80 },
            children: [
              new TextRun({
                text: `Honors: ${edu.honors}`,
                font: fontName,
                size: 18,
                italic: true
              })
            ]
          }));
        }
      });
    }

    // Skills Section
    if (resumeData.summary.skills?.length > 0) {
      addSectionHeader("Skills");
      children.push(new Paragraph({
        spacing: { before: 80, after: 120 },
        children: [
          new TextRun({
            text: resumeData.summary.skills.join("   •   "),
            font: fontName,
            size: 20
          })
        ]
      }));
    }

    // Certifications Section
    if (resumeData.certifications?.length > 0) {
      addSectionHeader("Certifications");
      resumeData.certifications.forEach(cert => {
        children.push(new Paragraph({
          tabStops: [{ type: "right", position: 10080 }],
          spacing: { before: 60, after: 60 },
          children: [
            new TextRun({
              text: `${cert.name}${cert.org ? `, ${cert.org}` : ''}`,
              font: fontName,
              size: 20
            }),
            new TextRun({
              text: `\t${cert.date}`,
              font: fontName,
              size: 20
            })
          ]
        }));
      });
    }

    // Projects Section
    if (resumeData.projects?.length > 0) {
      addSectionHeader("Projects");
      resumeData.projects.forEach(project => {
        children.push(new Paragraph({
          spacing: { before: 120, after: 40 },
          children: [
            new TextRun({
              text: project.name,
              font: fontName,
              size: 20,
              bold: true
            })
          ]
        }));

        if (project.description) {
          children.push(new Paragraph({
            spacing: { after: 80 },
            children: [
              new TextRun({
                text: project.description,
                font: fontName,
                size: 20,
                italic: true
              })
            ]
          }));
        }

        project.bullets.forEach(bullet => {
          if (bullet.trim()) {
            children.push(new Paragraph({
              bullet: { level: 0 },
              spacing: { after: 40 },
              children: [
                new TextRun({
                  text: bullet,
                  font: fontName,
                  size: 20
                })
              ]
            }));
          }
        });
      });
    }

    // Awards Section
    if (resumeData.awards?.length > 0) {
      addSectionHeader("Awards");
      resumeData.awards.forEach(award => {
        if (award.trim()) {
          children.push(new Paragraph({
            bullet: { level: 0 },
            spacing: { after: 60 },
            children: [
              new TextRun({
                text: award,
                font: fontName,
                size: 20
              })
            ]
          }));
        }
      });
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 1080, // 0.75in
              bottom: 1080,
              left: 1080,
              right: 1080
            }
          }
        },
        children: children
      }]
    });

    const filename = `${resumeData.name.replace(/\s+/g, '_')}_resume.docx`;
    const blob = await Packer.toBlob(doc);
    
    if (Capacitor.isNativePlatform()) {
      try {
        const base64Data = await blobToBase64(blob);
        const result = await Filesystem.writeFile({
          path: filename,
          data: base64Data,
          directory: Directory.Cache
        });

        await Share.share({
          title: 'Save Word Resume',
          url: result.uri
        });
      } catch (error) {
        console.error('Error generating or sharing Word doc on native:', error);
        alert('Failed to generate or share Word document: ' + error.message);
      }
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    }
  };

  // Font/Style selectors mapping based on chosen category and variant
  const getThemeClass = () => {
    switch (activeCategory) {
      case "corporate":
        return {
          font: "font-garamond",
          nameFont: "font-cinzel",
          accentColor: "#1B2B47",
          backgroundColor: "#fdfcf9",
          textColor: "text-slate-900"
        };
      case "tech":
        return {
          font: "font-inter",
          nameFont: "font-inter",
          accentColor: "#3B5168",
          backgroundColor: "#ffffff",
          textColor: "text-slate-900"
        };
      case "creative":
        return {
          font: "font-cardo",
          nameFont: "font-inter font-bold tracking-tight", // Space Grotesk layout style
          accentColor: creativeAccent,
          backgroundColor: "#fdfcf9",
          textColor: "text-slate-900"
        };
      case "healthcare":
        return {
          font: "font-lora",
          nameFont: "font-inter",
          accentColor: "#2A6B6E",
          backgroundColor: "#f8fafc", // Soft blue-grey
          textColor: "text-slate-900"
        };
      case "academic":
        return {
          font: "font-garamond",
          nameFont: "font-garamond",
          accentColor: "#6B2737",
          backgroundColor: "#ffffff",
          textColor: "text-slate-900"
        };
      case "trades":
        return {
          font: "font-sourceSans",
          nameFont: "font-sourceSans",
          accentColor: "#C2511F",
          backgroundColor: "#ffffff",
          textColor: "text-slate-900"
        };
      default:
        return {
          font: "font-garamond",
          nameFont: "font-cinzel",
          accentColor: "#1B2B47",
          backgroundColor: "#fdfcf9",
          textColor: "text-slate-900"
        };
    }
  };

  const theme = getThemeClass();

  return (
    <div className="flex flex-col min-h-screen bg-slate-950 text-slate-100 print-wrapper">
      
      {/* Top Header Controls (Hidden on Print) */}
      <header className="no-print bg-slate-900 border-b border-slate-800 py-4 px-6 flex flex-col sm:flex-row items-center justify-between gap-4 sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <FileText className="w-8 h-8 text-blue-500" />
          <div>
            <h1 className="text-lg font-bold tracking-tight bg-gradient-to-r from-blue-400 to-indigo-400 bg-clip-text text-transparent">
              ATS-Friendly Resume Builder
            </h1>
            <p className="text-xs text-slate-400">Create, switch templates, and export cleanly</p>
          </div>
        </div>

        {/* Try with Sample Resume & Keys */}
        <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto justify-end">
          <div className="relative">
            <input 
              type="password" 
              placeholder="Paste Anthropic API Key (Optional)" 
              value={apiKey} 
              onChange={(e) => setApiKey(e.target.value)}
              className="bg-slate-950 text-xs px-3 py-1.5 pr-8 rounded border border-slate-800 text-slate-300 focus:outline-none focus:border-blue-500 w-52"
            />
            <Settings className="w-3.5 h-3.5 absolute right-2.5 top-2.5 text-slate-500" />
          </div>

          <button
            onClick={handleTrySample}
            className="text-xs bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-1.5 rounded transition flex items-center gap-1.5 border border-slate-700"
          >
            <Sparkles className="w-3.5 h-3.5 text-amber-400" />
            Try with Sample
          </button>
        </div>
      </header>

      {/* Main Workspace Layout */}
      <main className="flex-1 flex flex-col lg:flex-row overflow-hidden print-wrapper">
        
        {/* Left Input & Editing Panel (Hidden on Print) */}
        <section className="no-print w-full lg:w-[450px] bg-slate-900 border-r border-slate-800 flex flex-col h-[calc(100vh-68px)] lg:h-[calc(100vh-68px)] overflow-y-auto p-5 shrink-0">
          
          {/* Main Raw Paste Textarea */}
          <div className="mb-6 bg-slate-950 p-4 rounded-lg border border-slate-800">
            <label className="block text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2 flex items-center justify-between">
              <span>Paste Raw Resume Text</span>
              <span className="text-[10px] text-slate-500 font-normal">AI parses sections instantly</span>
            </label>
            <textarea
              placeholder="Paste your existing resume here..."
              value={resumeText}
              onChange={(e) => setResumeText(e.target.value)}
              rows={6}
              className="w-full bg-slate-900 rounded border border-slate-800 p-2 text-xs font-mono text-slate-300 focus:outline-none focus:border-blue-500 resize-none"
            />
            <button
              onClick={handleParse}
              disabled={isParsing}
              className="w-full mt-3 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 disabled:from-slate-800 disabled:to-slate-800 text-white font-medium text-xs py-2 px-4 rounded transition flex items-center justify-center gap-2"
            >
              {isParsing ? (
                <>
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  Parsing resume...
                </>
              ) : (
                <>
                  <RefreshCw className="w-3.5 h-3.5" />
                  Parse Resume
                </>
              )}
            </button>
            {error && (
              <div className="mt-3 flex items-center gap-1.5 text-red-400 text-xs bg-red-950/30 p-2 rounded border border-red-900/30">
                <ShieldAlert className="w-4 h-4 shrink-0" />
                <span>{error}</span>
              </div>
            )}
          </div>

          <div className="flex items-center gap-1 mb-3 text-xs font-bold text-slate-400 uppercase tracking-wider">
            <Sliders className="w-3.5 h-3.5" />
            <span>Review & Edit Fields</span>
          </div>

          {/* Form Editing Accordion */}
          <div className="space-y-3 flex-1 pb-16">
            
            {/* Contact Details Accordion */}
            <div className="border border-slate-800 rounded-lg overflow-hidden">
              <button 
                onClick={() => toggleSection("contact")}
                className="w-full flex items-center justify-between bg-slate-950 p-3 text-left hover:bg-slate-800 transition"
              >
                <span className="text-xs font-bold uppercase tracking-wider text-slate-300">Contact Details</span>
                {openSections.contact ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {openSections.contact && (
                <div className="p-4 bg-slate-900/50 border-t border-slate-800 space-y-3">
                  <div>
                    <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Full Name</label>
                    <input 
                      type="text" 
                      value={resumeData.name} 
                      onChange={(e) => setResumeData(prev => ({ ...prev, name: e.target.value }))}
                      className="w-full bg-slate-950 text-xs px-2.5 py-1.5 rounded border border-slate-800 focus:outline-none focus:border-blue-500"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Professional Title</label>
                    <input 
                      type="text" 
                      value={resumeData.title} 
                      onChange={(e) => setResumeData(prev => ({ ...prev, title: e.target.value }))}
                      className="w-full bg-slate-950 text-xs px-2.5 py-1.5 rounded border border-slate-800 focus:outline-none focus:border-blue-500"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Phone</label>
                      <input 
                        type="text" 
                        value={resumeData.contact.phone} 
                        onChange={(e) => updateContact("phone", e.target.value)}
                        className="w-full bg-slate-950 text-xs px-2.5 py-1.5 rounded border border-slate-800"
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Email</label>
                      <input 
                        type="email" 
                        value={resumeData.contact.email} 
                        onChange={(e) => updateContact("email", e.target.value)}
                        className="w-full bg-slate-950 text-xs px-2.5 py-1.5 rounded border border-slate-800"
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Address / Location</label>
                    <input 
                      type="text" 
                      value={resumeData.contact.location} 
                      onChange={(e) => updateContact("location", e.target.value)}
                      className="w-full bg-slate-950 text-xs px-2.5 py-1.5 rounded border border-slate-800"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">LinkedIn</label>
                      <input 
                        type="text" 
                        value={resumeData.contact.linkedin} 
                        onChange={(e) => updateContact("linkedin", e.target.value)}
                        className="w-full bg-slate-950 text-xs px-2.5 py-1.5 rounded border border-slate-800"
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Website</label>
                      <input 
                        type="text" 
                        value={resumeData.contact.website} 
                        onChange={(e) => updateContact("website", e.target.value)}
                        className="w-full bg-slate-950 text-xs px-2.5 py-1.5 rounded border border-slate-800"
                      />
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Overview Section */}
            <div className="border border-slate-800 rounded-lg overflow-hidden">
              <button 
                onClick={() => toggleSection("summary")}
                className="w-full flex items-center justify-between bg-slate-950 p-3 text-left hover:bg-slate-800 transition"
              >
                <span className="text-xs font-bold uppercase tracking-wider text-slate-300">Professional Overview</span>
                {openSections.summary ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {openSections.summary && (
                <div className="p-4 bg-slate-900/50 border-t border-slate-800 space-y-3">
                  <div>
                    <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Tagline</label>
                    <textarea 
                      rows={3}
                      value={resumeData.summary.tagline} 
                      onChange={(e) => updateSummary("tagline", e.target.value)}
                      className="w-full bg-slate-950 text-xs p-2 rounded border border-slate-800 focus:outline-none focus:border-blue-500 resize-none"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Summary Bullet Points (One per line)</label>
                    <textarea 
                      rows={3}
                      value={resumeData.summary.bullets.join("\n")} 
                      onChange={(e) => updateSummary("bullets", e.target.value.split("\n"))}
                      className="w-full bg-slate-950 text-xs p-2 rounded border border-slate-800 font-mono focus:outline-none focus:border-blue-500 resize-none"
                    />
                  </div>
                </div>
              )}
            </div>

            {/* Work Experience repeating blocks */}
            <div className="border border-slate-800 rounded-lg overflow-hidden">
              <button 
                onClick={() => toggleSection("experience")}
                className="w-full flex items-center justify-between bg-slate-950 p-3 text-left hover:bg-slate-800 transition"
              >
                <span className="text-xs font-bold uppercase tracking-wider text-slate-300">Work Experience</span>
                {openSections.experience ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {openSections.experience && (
                <div className="p-4 bg-slate-900/50 border-t border-slate-800 space-y-4">
                  {resumeData.experience.map((entry, idx) => (
                    <div key={idx} className="bg-slate-950 p-3 rounded-lg border border-slate-800 relative space-y-3">
                      <button 
                        onClick={() => handleRemoveExperience(idx)}
                        className="absolute top-2 right-2 text-slate-500 hover:text-red-400 p-1 transition"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                      <h4 className="text-[10px] text-blue-400 uppercase font-bold">Role #{idx + 1}</h4>
                      
                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Company</label>
                          <input 
                            type="text" 
                            value={entry.company} 
                            onChange={(e) => handleUpdateExperience(idx, "company", e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        </div>
                        <div>
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Location</label>
                          <input 
                            type="text" 
                            value={entry.location} 
                            onChange={(e) => handleUpdateExperience(idx, "location", e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-3 gap-2">
                        <div className="col-span-2">
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Role Title</label>
                          <input 
                            type="text" 
                            value={entry.title} 
                            onChange={(e) => handleUpdateExperience(idx, "title", e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        </div>
                        <div>
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Dates</label>
                          <div className="flex gap-1">
                            <input 
                              type="text" 
                              placeholder="Start"
                              value={entry.start} 
                              onChange={(e) => handleUpdateExperience(idx, "start", e.target.value)}
                              className="w-full bg-slate-900 text-xs px-1.5 py-1 rounded border border-slate-800 text-center"
                            />
                            <input 
                              type="text" 
                              placeholder="End"
                              value={entry.end} 
                              onChange={(e) => handleUpdateExperience(idx, "end", e.target.value)}
                              className="w-full bg-slate-900 text-xs px-1.5 py-1 rounded border border-slate-800 text-center"
                            />
                          </div>
                        </div>
                      </div>

                      <div>
                        <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Overview Paragraph</label>
                        <input 
                          type="text" 
                          value={entry.summary} 
                          onChange={(e) => handleUpdateExperience(idx, "summary", e.target.value)}
                          className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                        />
                      </div>

                      <div className="space-y-1.5">
                        <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">3 Key Bullets</label>
                        {Array.from({ length: 3 }).map((_, bIdx) => (
                          <input 
                            key={bIdx}
                            type="text" 
                            placeholder={`Bullet point ${bIdx + 1}`}
                            value={entry.bullets[bIdx] || ""} 
                            onChange={(e) => handleUpdateExperienceBullet(idx, bIdx, e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        ))}
                      </div>
                    </div>
                  ))}
                  
                  <button
                    onClick={handleAddExperience}
                    className="w-full py-1.5 border border-dashed border-slate-800 hover:border-slate-600 rounded-lg text-xs text-slate-400 hover:text-slate-200 transition flex items-center justify-center gap-1.5"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    Add Work Experience
                  </button>
                </div>
              )}
            </div>

            {/* Education repeating blocks */}
            <div className="border border-slate-800 rounded-lg overflow-hidden">
              <button 
                onClick={() => toggleSection("education")}
                className="w-full flex items-center justify-between bg-slate-950 p-3 text-left hover:bg-slate-800 transition"
              >
                <span className="text-xs font-bold uppercase tracking-wider text-slate-300">Education</span>
                {openSections.education ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {openSections.education && (
                <div className="p-4 bg-slate-900/50 border-t border-slate-800 space-y-4">
                  {resumeData.education.map((edu, idx) => (
                    <div key={idx} className="bg-slate-950 p-3 rounded-lg border border-slate-800 relative space-y-3">
                      <button 
                        onClick={() => handleRemoveEducation(idx)}
                        className="absolute top-2 right-2 text-slate-500 hover:text-red-400 p-1 transition"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                      
                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">School</label>
                          <input 
                            type="text" 
                            value={edu.school} 
                            onChange={(e) => handleUpdateEducation(idx, "school", e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        </div>
                        <div>
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Location</label>
                          <input 
                            type="text" 
                            value={edu.location} 
                            onChange={(e) => handleUpdateEducation(idx, "location", e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-3 gap-2">
                        <div className="col-span-2">
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Degree</label>
                          <input 
                            type="text" 
                            value={edu.degree} 
                            onChange={(e) => handleUpdateEducation(idx, "degree", e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        </div>
                        <div>
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Graduation Year</label>
                          <input 
                            type="text" 
                            value={edu.date} 
                            onChange={(e) => handleUpdateEducation(idx, "date", e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800 text-center"
                          />
                        </div>
                      </div>

                      <div>
                        <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Honors (Optional)</label>
                        <input 
                          type="text" 
                          value={edu.honors} 
                          onChange={(e) => handleUpdateEducation(idx, "honors", e.target.value)}
                          className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                        />
                      </div>
                    </div>
                  ))}

                  <button
                    onClick={handleAddEducation}
                    className="w-full py-1.5 border border-dashed border-slate-800 hover:border-slate-600 rounded-lg text-xs text-slate-400 hover:text-slate-200 transition flex items-center justify-center gap-1.5"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    Add Education
                  </button>
                </div>
              )}
            </div>

            {/* Skills */}
            <div className="border border-slate-800 rounded-lg overflow-hidden">
              <button 
                onClick={() => toggleSection("skills")}
                className="w-full flex items-center justify-between bg-slate-950 p-3 text-left hover:bg-slate-800 transition"
              >
                <span className="text-xs font-bold uppercase tracking-wider text-slate-300">Skills</span>
                {openSections.skills ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {openSections.skills && (
                <div className="p-4 bg-slate-900/50 border-t border-slate-800 space-y-3">
                  <div>
                    <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Skills (Comma-separated list)</label>
                    <textarea 
                      rows={3}
                      value={resumeData.summary.skills.join(", ")} 
                      onChange={(e) => updateSummary("skills", e.target.value.split(",").map(s => s.trim()).filter(Boolean))}
                      className="w-full bg-slate-950 text-xs p-2.5 rounded border border-slate-800 focus:outline-none focus:border-blue-500 resize-none font-mono"
                    />
                  </div>
                </div>
              )}
            </div>

            {/* Certifications repeating blocks */}
            <div className="border border-slate-800 rounded-lg overflow-hidden">
              <button 
                onClick={() => toggleSection("certifications")}
                className="w-full flex items-center justify-between bg-slate-950 p-3 text-left hover:bg-slate-800 transition"
              >
                <span className="text-xs font-bold uppercase tracking-wider text-slate-300">Certifications</span>
                {openSections.certifications ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {openSections.certifications && (
                <div className="p-4 bg-slate-900/50 border-t border-slate-800 space-y-4">
                  {resumeData.certifications.map((cert, idx) => (
                    <div key={idx} className="bg-slate-950 p-3 rounded-lg border border-slate-800 relative space-y-3">
                      <button 
                        onClick={() => handleRemoveCertification(idx)}
                        className="absolute top-2 right-2 text-slate-500 hover:text-red-400 p-1 transition"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                      
                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Certification Name</label>
                          <input 
                            type="text" 
                            value={cert.name} 
                            onChange={(e) => handleUpdateCertification(idx, "name", e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        </div>
                        <div>
                          <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Organization</label>
                          <input 
                            type="text" 
                            value={cert.org} 
                            onChange={(e) => handleUpdateCertification(idx, "org", e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        </div>
                      </div>
                      <div>
                        <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Date</label>
                        <input 
                          type="text" 
                          value={cert.date} 
                          onChange={(e) => handleUpdateCertification(idx, "date", e.target.value)}
                          className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                        />
                      </div>
                    </div>
                  ))}

                  <button
                    onClick={handleAddCertification}
                    className="w-full py-1.5 border border-dashed border-slate-800 hover:border-slate-600 rounded-lg text-xs text-slate-400 hover:text-slate-200 transition flex items-center justify-center gap-1.5"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    Add Certification
                  </button>
                </div>
              )}
            </div>

            {/* Projects repeating blocks */}
            <div className="border border-slate-800 rounded-lg overflow-hidden">
              <button 
                onClick={() => toggleSection("projects")}
                className="w-full flex items-center justify-between bg-slate-950 p-3 text-left hover:bg-slate-800 transition"
              >
                <span className="text-xs font-bold uppercase tracking-wider text-slate-300">Projects</span>
                {openSections.projects ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {openSections.projects && (
                <div className="p-4 bg-slate-900/50 border-t border-slate-800 space-y-4">
                  {resumeData.projects.map((proj, idx) => (
                    <div key={idx} className="bg-slate-950 p-3 rounded-lg border border-slate-800 relative space-y-3">
                      <button 
                        onClick={() => handleRemoveProject(idx)}
                        className="absolute top-2 right-2 text-slate-500 hover:text-red-400 p-1 transition"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                      
                      <div>
                        <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Project Name</label>
                        <input 
                          type="text" 
                          value={proj.name} 
                          onChange={(e) => handleUpdateProject(idx, "name", e.target.value)}
                          className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                        />
                      </div>
                      <div>
                        <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Description</label>
                        <input 
                          type="text" 
                          value={proj.description} 
                          onChange={(e) => handleUpdateProject(idx, "description", e.target.value)}
                          className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="block text-[9px] text-slate-500 uppercase font-bold mb-1">Bullets</label>
                        {Array.from({ length: 2 }).map((_, bIdx) => (
                          <input 
                            key={bIdx}
                            type="text" 
                            value={proj.bullets[bIdx] || ""} 
                            onChange={(e) => handleUpdateProjectBullet(idx, bIdx, e.target.value)}
                            className="w-full bg-slate-900 text-xs px-2 py-1 rounded border border-slate-800"
                          />
                        ))}
                      </div>
                    </div>
                  ))}

                  <button
                    onClick={handleAddProject}
                    className="w-full py-1.5 border border-dashed border-slate-800 hover:border-slate-600 rounded-lg text-xs text-slate-400 hover:text-slate-200 transition flex items-center justify-center gap-1.5"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    Add Project
                  </button>
                </div>
              )}
            </div>

            {/* Awards */}
            <div className="border border-slate-800 rounded-lg overflow-hidden">
              <button 
                onClick={() => toggleSection("awards")}
                className="w-full flex items-center justify-between bg-slate-950 p-3 text-left hover:bg-slate-800 transition"
              >
                <span className="text-xs font-bold uppercase tracking-wider text-slate-300">Awards</span>
                {openSections.awards ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {openSections.awards && (
                <div className="p-4 bg-slate-900/50 border-t border-slate-800 space-y-3">
                  <div>
                    <label className="block text-[10px] text-slate-500 uppercase font-bold mb-1">Awards (One per line)</label>
                    <textarea 
                      rows={3}
                      value={resumeData.awards.join("\n")} 
                      onChange={(e) => setResumeData(prev => ({ ...prev, awards: e.target.value.split("\n").filter(Boolean) }))}
                      className="w-full bg-slate-950 text-xs p-2.5 rounded border border-slate-800 focus:outline-none focus:border-blue-500 resize-none font-mono"
                    />
                  </div>
                </div>
              )}
            </div>

          </div>
        </section>

        {/* Right Live Preview Panel */}
        <section className="flex-1 flex flex-col h-[calc(100vh-68px)] overflow-y-auto bg-slate-950 p-6 items-center print-container">
          
          {/* Template Category Tabs (Hidden on Print) */}
          <div className="no-print w-full max-w-[8.5in] mb-6 space-y-4">
            <div className="bg-slate-900 p-1.5 rounded-lg border border-slate-800 flex flex-wrap gap-1">
              {[
                { id: "corporate", name: "Corporate" },
                { id: "tech", name: "Tech / Eng" },
                { id: "creative", name: "Creative" },
                { id: "healthcare", name: "Healthcare" },
                { id: "academic", name: "Academic" },
                { id: "trades", name: "Trades" }
              ].map(cat => (
                <button
                  key={cat.id}
                  onClick={() => selectCategory(cat.id)}
                  className={`flex-1 min-w-[90px] text-xs font-semibold py-1.5 px-3 rounded transition ${
                    activeCategory === cat.id 
                      ? 'bg-blue-600 text-white shadow-sm' 
                      : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
                  }`}
                >
                  {cat.name}
                </button>
              ))}
            </div>

            {/* Template Variant Selectors & Theme Settings */}
            <div className="flex flex-wrap items-center justify-between gap-4 bg-slate-900/40 p-3 rounded-lg border border-slate-800/60">
              <div className="flex items-center gap-2">
                <span className="text-xs text-slate-400 font-medium">Variant:</span>
                
                {activeCategory === "corporate" && (
                  <div className="flex gap-1.5">
                    <button 
                      onClick={() => setActiveVariant("boardroom")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'boardroom' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Boardroom (A)
                    </button>
                    <button 
                      onClick={() => setActiveVariant("partner")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'partner' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Partner (B)
                    </button>
                  </div>
                )}

                {activeCategory === "tech" && (
                  <div className="flex gap-1.5">
                    <button 
                      onClick={() => setActiveVariant("builder")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'builder' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Builder (A)
                    </button>
                    <button 
                      onClick={() => setActiveVariant("architect")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'architect' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Architect (B)
                    </button>
                  </div>
                )}

                {activeCategory === "creative" && (
                  <div className="flex gap-1.5">
                    <button 
                      onClick={() => setActiveVariant("editorial")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'editorial' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Editorial (A)
                    </button>
                    <button 
                      onClick={() => setActiveVariant("studio")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'studio' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Studio (B)
                    </button>
                  </div>
                )}

                {activeCategory === "healthcare" && (
                  <div className="flex gap-1.5">
                    <button 
                      onClick={() => setActiveVariant("practitioner")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'practitioner' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Practitioner (A)
                    </button>
                    <button 
                      onClick={() => setActiveVariant("educator")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'educator' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Educator (B)
                    </button>
                  </div>
                )}

                {activeCategory === "academic" && (
                  <div className="flex gap-1.5">
                    <button 
                      onClick={() => setActiveVariant("scholar")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'scholar' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Scholar (A)
                    </button>
                    <button 
                      onClick={() => setActiveVariant("researcher")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'researcher' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Researcher (B)
                    </button>
                  </div>
                )}

                {activeCategory === "trades" && (
                  <div className="flex gap-1.5">
                    <button 
                      onClick={() => setActiveVariant("operator")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'operator' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Operator (A)
                    </button>
                    <button 
                      onClick={() => setActiveVariant("foreman")}
                      className={`text-xs px-3 py-1 rounded border transition ${activeVariant === 'foreman' ? 'bg-slate-800 border-blue-500 text-blue-400' : 'bg-slate-950 border-slate-800 text-slate-400'}`}
                    >
                      The Foreman (B)
                    </button>
                  </div>
                )}
              </div>

              {/* Creative Theme Color Toggle */}
              {activeCategory === "creative" && (
                <div className="flex items-center gap-2 border-l border-slate-800 pl-4">
                  <span className="text-xs text-slate-400 font-medium">Color:</span>
                  <button
                    onClick={() => setCreativeAccent("#A85432")}
                    className={`w-4 h-4 rounded-full border transition-transform ${creativeAccent === '#A85432' ? 'border-white scale-125' : 'border-transparent'}`}
                    style={{ backgroundColor: "#A85432" }}
                    title="Clay Accent"
                  />
                  <button
                    onClick={() => setCreativeAccent("#6B7F5C")}
                    className={`w-4 h-4 rounded-full border transition-transform ${creativeAccent === '#6B7F5C' ? 'border-white scale-125' : 'border-transparent'}`}
                    style={{ backgroundColor: "#6B7F5C" }}
                    title="Sage Accent"
                  />
                </div>
              )}
            </div>
          </div>

          {/* Letter sized Page Frame */}
          <div className="relative w-full max-w-[8.5in] print-container mb-24">
            
            {/* Page Break lines (On-screen overlay, hidden on print) */}
            {pageBreaks.map((breakY, idx) => (
              <div 
                key={idx} 
                className="page-break-indicator no-print" 
                style={{ top: `${breakY}px` }} 
              />
            ))}

            {/* Actual Resume Content Div */}
            <div
              ref={previewRef}
              id="resume-pdf-content"
              className={`resume-page resume-page-shadow rounded-sm mx-auto relative ${theme.textColor} ${theme.font}`}
              style={{
                width: '8.5in',
                minHeight: '11in',
                padding: '0.75in',
                backgroundColor: theme.backgroundColor,
                color: '#1a1a1a', // Rich dark print tone
                transition: 'all 0.2s ease-in-out'
              }}
            >
              
              {/* Variant-specific decorative marks */}
              {activeCategory === "corporate" && activeVariant === "boardroom" && (
                // Navy quarter-circle in top-right corner
                <div 
                  className="absolute top-0 right-0 w-16 h-16 rounded-bl-full no-print" 
                  style={{ backgroundColor: theme.accentColor }} 
                />
              )}

              {activeCategory === "trades" && activeVariant === "foreman" && (
                // Charcoal quarter-circle in top-right corner
                <div 
                  className="absolute top-0 right-0 w-16 h-16 rounded-bl-full no-print" 
                  style={{ backgroundColor: "#2D2D2D" }} 
                />
              )}

              {/* Asymmetric layout accent sidebar (Creative Editorial Variant A) */}
              {activeCategory === "creative" && activeVariant === "editorial" && (
                <div 
                  className="absolute top-0 left-0 bottom-0 w-1.5"
                  style={{ backgroundColor: theme.accentColor }}
                />
              )}

              {/* Layout structures */}
              
              {/* Header block */}
              {activeCategory === "creative" && activeVariant === "editorial" ? (
                // Asymmetric Left-Right Name and Contact on same baseline
                <div className="flex flex-col md:flex-row justify-between items-start md:items-baseline border-b pb-4 mb-6" style={{ borderColor: `${theme.accentColor}40` }}>
                  <div>
                    <h1 className="text-3xl font-extrabold tracking-tight uppercase" style={{ color: theme.accentColor }}>{resumeData.name}</h1>
                    <p className="text-xs uppercase tracking-widest font-semibold text-slate-500 mt-1">{resumeData.title}</p>
                  </div>
                  <div className="text-right text-[10px] space-y-0.5 mt-2 md:mt-0 font-medium text-slate-600">
                    <div>{resumeData.contact.phone} • {resumeData.contact.email}</div>
                    <div>{resumeData.contact.location}</div>
                    <div>{resumeData.contact.linkedin} {resumeData.contact.website && `• ${resumeData.contact.website}`}</div>
                  </div>
                </div>
              ) : activeCategory === "academic" && activeVariant === "researcher" ? (
                // academic compact layout header
                <div className="border-b pb-3 mb-4 text-center">
                  <h1 className="text-2xl font-bold uppercase tracking-wider">{resumeData.name}</h1>
                  <p className="text-xs italic text-slate-600 mt-0.5">{resumeData.title}</p>
                  <div className="text-[10px] text-slate-600 mt-2 flex flex-wrap justify-center gap-x-3 gap-y-1">
                    <span>{resumeData.contact.phone}</span>
                    <span>•</span>
                    <span>{resumeData.contact.email}</span>
                    <span>•</span>
                    <span>{resumeData.contact.location}</span>
                    {resumeData.contact.linkedin && (
                      <>
                        <span>•</span>
                        <span>{resumeData.contact.linkedin}</span>
                      </>
                    )}
                    {resumeData.contact.website && (
                      <>
                        <span>•</span>
                        <span>{resumeData.contact.website}</span>
                      </>
                    )}
                  </div>
                </div>
              ) : (
                // Standard centered header
                <div className="text-center mb-6">
                  {/* Name */}
                  <h1 
                    className={`text-3xl font-bold uppercase ${theme.nameFont}`}
                    style={{ 
                      color: activeCategory === "trades" && activeVariant === "operator" ? "#2D2D2D" : theme.accentColor,
                      letterSpacing: activeCategory === "corporate" ? "0.3em" : "0.05em"
                    }}
                  >
                    {resumeData.name}
                  </h1>

                  {/* Title */}
                  {resumeData.title && (
                    <p 
                      className={`text-xs uppercase tracking-widest mt-1 text-slate-500 font-semibold`}
                    >
                      {resumeData.title}
                    </p>
                  )}

                  {/* Single horizontal rule separator */}
                  <hr className="my-3 border-t" style={{ borderColor: `${theme.accentColor}30` }} />

                  {/* Contact Row (Pipe-separated) */}
                  <div className="text-[10px] text-slate-600 flex flex-wrap justify-center items-center gap-x-3 gap-y-1 font-medium">
                    <span>{resumeData.contact.phone}</span>
                    {resumeData.contact.email && <span className="opacity-40">|</span>}
                    <span>{resumeData.contact.email}</span>
                    {resumeData.contact.location && <span className="opacity-40">|</span>}
                    <span>{resumeData.contact.location}</span>
                    {resumeData.contact.linkedin && <span className="opacity-40">|</span>}
                    <span>{resumeData.contact.linkedin}</span>
                    {resumeData.contact.website && <span className="opacity-40">|</span>}
                    <span>{resumeData.contact.website}</span>
                  </div>
                </div>
              )}

              {/* Sub-components/Section Headers */}
              {/* We define a reusable renderer that respects categories */}
              {(() => {
                const renderSectionHeader = (title) => {
                  if (activeCategory === "corporate" && activeVariant === "boardroom") {
                    return (
                      <div className="w-full bg-[#E8E8E8] py-1 px-3 mb-3 text-center border-b border-t border-slate-300">
                        <h2 className="text-xs uppercase font-bold tracking-widest text-[#1B2B47]">{title}</h2>
                      </div>
                    );
                  }
                  
                  if (activeCategory === "corporate" && activeVariant === "partner") {
                    return (
                      <div className="w-full border-b border-[#1B2B47] pb-1 mb-3 text-left">
                        <h2 className="text-xs uppercase font-bold tracking-wider text-[#1B2B47]" style={{ fontVariant: 'all-small-caps' }}>{title}</h2>
                      </div>
                    );
                  }

                  if (activeCategory === "tech" && activeVariant === "builder") {
                    return (
                      <div className="w-full border-b border-[#3B5168] pb-0.5 mb-3 text-left">
                        <h2 className="text-xs uppercase font-bold tracking-wide text-[#3B5168]">{title}</h2>
                      </div>
                    );
                  }

                  if (activeCategory === "tech" && activeVariant === "architect") {
                    return (
                      <div className="w-full bg-[#3B5168]/10 py-1 px-2 mb-3 text-left border-l-4 border-[#3B5168]">
                        <h2 className="text-xs uppercase font-bold tracking-wide text-[#3B5168]">{title}</h2>
                      </div>
                    );
                  }

                  if (activeCategory === "creative" && activeVariant === "editorial") {
                    return (
                      <div className="w-full border-b border-accent pb-0.5 mb-3 text-left" style={{ borderColor: `${theme.accentColor}30` }}>
                        <h2 className="text-xs uppercase font-bold tracking-wider italic text-slate-800" style={{ color: theme.accentColor, fontVariant: 'all-small-caps' }}>{title}</h2>
                      </div>
                    );
                  }

                  if (activeCategory === "creative" && activeVariant === "studio") {
                    return (
                      <div className="w-full py-1 px-3 mb-3 text-center rounded" style={{ backgroundColor: `${theme.accentColor}15` }}>
                        <h2 className="text-xs uppercase font-bold tracking-widest" style={{ color: theme.accentColor }}>{title}</h2>
                      </div>
                    );
                  }

                  if (activeCategory === "healthcare" && activeVariant === "practitioner") {
                    return (
                      <div className="w-full bg-[#2A6B6E] text-white py-1 px-3 mb-3 text-center font-semibold rounded-sm">
                        <h2 className="text-xs uppercase font-bold tracking-wider">{title}</h2>
                      </div>
                    );
                  }

                  if (activeCategory === "healthcare" && activeVariant === "educator") {
                    return (
                      <div className="w-full border-b-2 border-[#2A6B6E] pb-1 mb-3 text-left">
                        <h2 className="text-xs uppercase font-bold tracking-wider text-[#2A6B6E]">{title}</h2>
                      </div>
                    );
                  }

                  if (activeCategory === "academic" && activeVariant === "scholar") {
                    return (
                      <div className="w-full border-b border-[#6B2737] pb-1 mb-3 text-center">
                        <h2 className="text-xs uppercase font-bold tracking-widest text-[#6B2737]">{title}</h2>
                      </div>
                    );
                  }

                  if (activeCategory === "academic" && activeVariant === "researcher") {
                    return (
                      <div className="w-full border-b border-[#6B2737] pb-0.5 mb-2.5 text-left">
                        <h2 className="text-[10px] uppercase font-bold tracking-wide text-[#6B2737]" style={{ fontVariant: 'all-small-caps' }}>{title}</h2>
                      </div>
                    );
                  }

                  if (activeCategory === "trades" && activeVariant === "operator") {
                    return (
                      <div className="w-full bg-[#2D2D2D] text-white py-1.5 px-3 mb-3 text-left flex justify-between items-center">
                        <h2 className="text-xs uppercase font-bold tracking-wide">{title}</h2>
                        <div className="w-12 h-1 bg-[#C2511F]" />
                      </div>
                    );
                  }

                  if (activeCategory === "trades" && activeVariant === "foreman") {
                    return (
                      <div className="w-full border-b border-[#2D2D2D] pb-1 mb-3 text-center">
                        <h2 className="text-xs uppercase font-bold tracking-widest text-[#2D2D2D]">{title}</h2>
                      </div>
                    );
                  }

                  return (
                    <div className="w-full border-b pb-1 mb-3 text-left">
                      <h2 className="text-xs uppercase font-bold tracking-wider">{title}</h2>
                    </div>
                  );
                };

                // Certifications & Licenses layout helper (for Foreman, Practitioner, etc.)
                const renderCertifications = () => {
                  if (resumeData.certifications?.length === 0) return null;
                  return (
                    <div className="resume-section-item mb-5">
                      {renderSectionHeader("Certifications & Licenses")}
                      <div className="grid grid-cols-2 gap-4">
                        {resumeData.certifications.map((cert, idx) => (
                          <div key={idx} className="text-[11px] leading-relaxed">
                            <span className="font-bold text-slate-800">{cert.name}</span>
                            {cert.org && <span className="text-slate-500"> — {cert.org}</span>}
                            {cert.date && <span className="float-right font-medium text-slate-600">{cert.date}</span>}
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                };

                // Work Experience list helper
                const renderExperience = () => {
                  if (resumeData.experience?.length === 0) return null;
                  return (
                    <div className="resume-section-item mb-5">
                      {renderSectionHeader("Work Experience")}
                      <div className="space-y-4">
                        {resumeData.experience.map((entry, idx) => (
                          <div key={idx} className="text-[11px] leading-relaxed">
                            {/* Company | Location on Left, Date on Right */}
                            <div className="flex justify-between items-baseline font-bold text-slate-900">
                              <span>{entry.company} {entry.location && `| ${entry.location}`}</span>
                              <span className="font-medium text-slate-700">{entry.start} – {entry.end}</span>
                            </div>

                            {/* Title on left */}
                            <div className="flex justify-between items-baseline text-slate-700 italic mt-0.5">
                              <span>{entry.title}</span>
                            </div>

                            {/* Optional role-specific tagline (Creative Studio) */}
                            {activeCategory === "creative" && activeVariant === "studio" && entry.summary && (
                              <div className="text-[10px] text-slate-500 italic mt-1 pl-2 border-l border-accent/40" style={{ borderColor: theme.accentColor }}>
                                {entry.summary}
                              </div>
                            )}

                            {/* Generic Overview Paragraph (if not Studio) */}
                            {!(activeCategory === "creative" && activeVariant === "studio") && entry.summary && (
                              <p className="text-slate-600 mt-1 text-justify">{entry.summary}</p>
                            )}

                            {/* Bullet points */}
                            <ul className="mt-1.5 space-y-1">
                              {entry.bullets.map((bullet, bIdx) => {
                                if (!bullet.trim()) return null;
                                return (
                                  <li key={bIdx} className="text-slate-600 pl-4 relative text-justify">
                                    <span className="absolute left-0">
                                      {activeCategory === "tech" && activeVariant === "builder" ? "→" : "—"}
                                    </span>
                                    {bullet}
                                  </li>
                                );
                              })}
                            </ul>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                };

                // Render elements in order depending on template options
                return (
                  <div className="space-y-5">
                    
                    {/* Summary Section */}
                    {(resumeData.summary.tagline || resumeData.summary.bullets?.length > 0) && (
                      <div className="resume-section-item mb-5">
                        {renderSectionHeader("Professional Overview")}
                        <div className="text-[11px] leading-relaxed">
                          {resumeData.summary.tagline && (
                            <p className="italic text-slate-700 font-medium mb-1.5 text-justify">{resumeData.summary.tagline}</p>
                          )}
                          {resumeData.summary.bullets?.length > 0 && (
                            <ul className="space-y-1">
                              {resumeData.summary.bullets.map((bullet, idx) => {
                                if (!bullet.trim()) return null;
                                return (
                                  <li key={idx} className="text-slate-600 pl-4 relative text-justify">
                                    <span className="absolute left-0">—</span>
                                    {bullet}
                                  </li>
                                );
                              })}
                            </ul>
                          )}
                        </div>
                      </div>
                    )}

                    {/* FOREMAN: Certifications prominently ABOVE experience */}
                    {activeCategory === "trades" && activeVariant === "foreman" && renderCertifications()}

                    {/* Work Experience */}
                    {renderExperience()}

                    {/* NOT FOREMAN: Certifications normally placed */}
                    {!(activeCategory === "trades" && activeVariant === "foreman") && renderCertifications()}

                    {/* Education Section */}
                    {resumeData.education?.length > 0 && (
                      <div className="resume-section-item mb-5">
                        {renderSectionHeader("Education")}
                        <div className="space-y-3">
                          {resumeData.education.map((edu, idx) => (
                            <div key={idx} className="text-[11px] leading-relaxed">
                              <div className="flex justify-between items-baseline font-bold text-slate-900">
                                <span>{edu.school} {edu.location && `| ${edu.location}`}</span>
                                <span className="font-medium text-slate-700">{edu.date}</span>
                              </div>
                              <div className="text-slate-700 mt-0.5">{edu.degree}</div>
                              {edu.honors && (
                                <div className="text-slate-500 italic text-[10px] mt-0.5">Honors: {edu.honors}</div>
                              )}
                            </div>
                          ))}
                        </div>
                      </div>
                    )}

                    {/* Skills Section */}
                    {resumeData.summary.skills?.length > 0 && (
                      <div className="resume-section-item mb-5">
                        {renderSectionHeader("Skills")}
                        
                        {/* Creative Theme tag badges */}
                        {activeCategory === "creative" && activeVariant === "studio" ? (
                          <div className="flex flex-wrap gap-1.5 mt-1.5">
                            {resumeData.summary.skills.map((skill, idx) => (
                              <span 
                                key={idx} 
                                className="text-[9px] px-2 py-0.5 rounded font-medium border text-slate-700"
                                style={{ 
                                  borderColor: `${theme.accentColor}40`,
                                  backgroundColor: `${theme.accentColor}08`
                                }}
                              >
                                {skill}
                              </span>
                            ))}
                          </div>
                        ) : activeCategory === "tech" && activeVariant === "builder" ? (
                          // Tech / Engineer Builder: Monospace accent blocks
                          <div className="flex flex-wrap gap-1.5 mt-1.5 font-mono">
                            {resumeData.summary.skills.map((skill, idx) => (
                              <span 
                                key={idx} 
                                className="text-[9px] bg-slate-100 px-2 py-0.5 rounded border border-slate-200 text-slate-800"
                              >
                                {skill}
                              </span>
                            ))}
                          </div>
                        ) : activeCategory === "tech" && activeVariant === "architect" ? (
                          // Tech Architect: Text-based inline pills
                          <div className="flex flex-wrap gap-1 mt-1 font-mono">
                            {resumeData.summary.skills.map((skill, idx) => (
                              <span 
                                key={idx} 
                                className="text-[9px] border border-slate-300 rounded px-1.5 py-0.5 inline-block text-slate-700"
                              >
                                {skill}
                              </span>
                            ))}
                          </div>
                        ) : activeCategory === "trades" && activeVariant === "operator" ? (
                          // Trades: plain dash-separated rows
                          <div className="text-[11px] text-slate-700 leading-relaxed mt-1 flex flex-wrap gap-x-3 gap-y-1">
                            {resumeData.summary.skills.map((skill, idx) => (
                              <span key={idx} className="whitespace-nowrap">
                                {skill}{idx < resumeData.summary.skills.length - 1 && "  —"}
                              </span>
                            ))}
                          </div>
                        ) : (
                          // Corporate & standard: 4-column grid
                          <div className="grid grid-cols-2 sm:grid-cols-4 gap-x-4 gap-y-1.5 text-[11px] text-slate-700 mt-1.5">
                            {resumeData.summary.skills.map((skill, idx) => (
                              <div key={idx} className="flex items-center gap-1.5">
                                <span className="text-[8px] opacity-40">•</span>
                                <span>{skill}</span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    )}

                    {/* Academic / Educator Publications Section */}
                    {((activeCategory === "academic") || (activeCategory === "healthcare" && activeVariant === "educator")) && (
                      <>
                        {/* Scholar / Researcher Publications */}
                        <div className="resume-section-item mb-5">
                          {renderSectionHeader("Publications & Presentations")}
                          <div className="text-[11px] leading-relaxed space-y-2">
                            {resumeData.projects?.map((proj, idx) => (
                              <div key={idx} className="pl-4 relative">
                                <span className="absolute left-0">•</span>
                                <span className="font-bold text-slate-800">{proj.name}</span>
                                {proj.description && <span className="italic"> — {proj.description}</span>}
                                {proj.bullets?.length > 0 && (
                                  <ul className="mt-0.5 space-y-0.5 text-slate-600 pl-4 list-circle">
                                    {proj.bullets.map((b, bIdx) => b.trim() && <li key={bIdx}>{b}</li>)}
                                  </ul>
                                )}
                              </div>
                            ))}
                          </div>
                        </div>

                        {/* Scholar / Academic Grants & Teaching Section */}
                        {activeCategory === "academic" && activeVariant === "scholar" && (
                          <div className="resume-section-item mb-5">
                            {renderSectionHeader("Grants & Academic Honors")}
                            <ul className="text-[11px] leading-relaxed space-y-1 pl-4 list-disc text-slate-600">
                              {resumeData.awards.map((award, idx) => (
                                <li key={idx}>{award}</li>
                              ))}
                            </ul>
                          </div>
                        )}
                      </>
                    )}

                    {/* Projects Section (for Tech, Creative, Trades, etc.) */}
                    {!(activeCategory === "academic" || (activeCategory === "healthcare" && activeVariant === "educator")) && resumeData.projects?.length > 0 && (
                      <div className="resume-section-item mb-5">
                        {renderSectionHeader("Projects")}
                        <div className="space-y-3">
                          {resumeData.projects.map((project, idx) => (
                            <div key={idx} className="text-[11px] leading-relaxed">
                              <div className="font-bold text-slate-900">{project.name}</div>
                              {project.description && (
                                <p className="text-slate-600 italic mt-0.5">{project.description}</p>
                              )}
                              {project.bullets?.length > 0 && (
                                <ul className="mt-1 space-y-1">
                                  {project.bullets.map((bullet, bIdx) => {
                                    if (!bullet.trim()) return null;
                                    return (
                                      <li key={bIdx} className="text-slate-600 pl-4 relative">
                                        <span className="absolute left-0">—</span>
                                        {bullet}
                                      </li>
                                    );
                                  })}
                                </ul>
                              )}
                            </div>
                          ))}
                        </div>
                      </div>
                    )}

                    {/* Awards Section (for general layout) */}
                    {!(activeCategory === "academic" && activeVariant === "scholar") && resumeData.awards?.length > 0 && (
                      <div className="resume-section-item mb-5">
                        {renderSectionHeader("Honors & Awards")}
                        <ul className="text-[11px] leading-relaxed space-y-1 text-slate-600 pl-4 list-disc">
                          {resumeData.awards.map((award, idx) => (
                            <li key={idx}>{award}</li>
                          ))}
                        </ul>
                      </div>
                    )}

                  </div>
                );
              })()}

            </div>
          </div>
        </section>

      </main>

      {/* Sticky Bottom Export Bar (Hidden on Print) */}
      <div className="no-print fixed bottom-6 right-6 bg-slate-900/90 backdrop-blur border border-slate-800 p-3 rounded-xl shadow-2xl flex items-center gap-3 z-50">
        
        {/* Native Print PDF Button */}
        <button
          onClick={() => window.print()}
          className="bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-xs py-2.5 px-4 rounded-lg transition flex items-center gap-2 shadow-lg"
          title="Print cleanly via the browser engine (produces selectable text)"
        >
          <Eye className="w-4 h-4" />
          Print PDF
        </button>

        {/* jsPDF/html2pdf Download PDF Button */}
        <button
          onClick={handleDownloadPDF}
          className="bg-blue-600 hover:bg-blue-500 text-white font-semibold text-xs py-2.5 px-4 rounded-lg transition flex items-center gap-2 shadow-lg"
          title="Download PDF directly in-browser using html2pdf.js"
        >
          <Download className="w-4 h-4" />
          Download PDF
        </button>

        {/* Word docx Button */}
        <button
          onClick={handleDownloadWord}
          className="bg-indigo-600 hover:bg-indigo-500 text-white font-semibold text-xs py-2.5 px-4 rounded-lg transition flex items-center gap-2 shadow-lg"
          title="Download Microsoft Word .docx file"
        >
          <Download className="w-4 h-4" />
          Download Word
        </button>
      </div>

    </div>
  );
}
