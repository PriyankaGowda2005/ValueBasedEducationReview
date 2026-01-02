"""
Script to create PowerPoint presentation based on LaTeX document
Value_Based_Education_Review_Overleaf.tex
Requires: python-pptx library (install with: pip install python-pptx)
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def create_presentation():
    # Create presentation object
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Define slide layouts
    title_slide_layout = prs.slide_layouts[0]  # Title slide
    content_slide_layout = prs.slide_layouts[1]  # Title and content
    blank_layout = prs.slide_layouts[6]  # Blank layout
    
    # SLIDE 1: TITLE SLIDE
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = "Scenario-Based Analytical Review of Value-Based Education and Student Mental Well-Being Using Traditional Indian Texts"
    subtitle.text = "A Narrative Literature Review\n\nDivya\nJanuary 2026\n\n[This is a narrative literature review - not an empirical study]"
    
    # SLIDE 2: DISCLOSURE STATEMENT
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "CRITICAL POSITIONING"
    tf = content.text_frame
    tf.text = "• This presentation presents a NARRATIVE LITERATURE REVIEW"
    p = tf.add_paragraph()
    p.text = "• NOT an empirical intervention study"
    p = tf.add_paragraph()
    p.text = "• NO claims of statistical improvement or effectiveness"
    p = tf.add_paragraph()
    p.text = "• Scenario-based analysis = conceptual exploration, NOT case studies"
    p = tf.add_paragraph()
    p.text = "• All findings are positioned as CONCEPTUAL CONTRIBUTIONS"
    p = tf.add_paragraph()
    p.text = "• Explicit limitations acknowledged throughout"
    
    # SLIDE 3: TABLE OF CONTENTS
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Table of Contents"
    tf = content.text_frame
    tf.text = "PRESENTATION OUTLINE"
    p = tf.add_paragraph()
    p.text = "1. Background & Problem Statement"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "2. Aim & Objectives"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "3. Conceptual Framework"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "4. Traditional Indian Texts"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "   • Basavanna's Vachanas"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "   • Panchatantra"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "   • Hitopadesha"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "5. Ancient Education Methods"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "6. Methodology"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "7. Rationale for Scenario-Based Analysis"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "8. Scenario-Based Analytical Review (10 Scenarios)"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "9. Synthesis of Findings"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "10. Critical Limitations"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "11. Future Research Directions"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "12. Conclusion"
    p.level = 1
    
    # SLIDE 4: BACKGROUND & PROBLEM STATEMENT
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Background & Problem Statement"
    tf = content.text_frame
    tf.text = "THE CONTEMPORARY EDUCATIONAL CHALLENGE"
    p = tf.add_paragraph()
    p.text = "• Rising mental health challenges among school students globally"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Stress, anxiety, emotional instability (WHO, 2021)"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Academic pressure and performance anxiety"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Peer conflicts and social isolation"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "• Overemphasis on academic performance"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Cognitive development prioritized over holistic growth"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Limited integration of values and emotional intelligence"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "• Fragmented interventions"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Lack of culturally grounded frameworks"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Need for comprehensive value-based approaches"
    p.level = 2
    
    # SLIDE 5: AIM & OBJECTIVES
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Aim & Objectives"
    tf = content.text_frame
    tf.text = "Primary Aim:"
    p = tf.add_paragraph()
    p.text = "To critically review value-based education literature and analyze traditional Indian texts as conceptual frameworks for mental health support"
    p = tf.add_paragraph()
    p.text = "Specific Objectives:"
    p = tf.add_paragraph()
    p.text = "• Synthesize literature on value-based education (2015-2025)"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Analyze traditional Indian texts as potential conceptual frameworks"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Employ scenario-based analytical approach"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Identify limitations and research gaps"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Position work as conceptual contribution requiring empirical validation"
    p.level = 1
    
    # SLIDE 6: CONCEPTUAL FRAMEWORK
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Conceptual Framework - Value-Based Education"
    tf = content.text_frame
    tf.text = "Value-Based Education:"
    p = tf.add_paragraph()
    p.text = "• Integrates moral, ethical, and life skills development"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Addresses multiple dimensions: cognitive, emotional, social, moral"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Recognizes interconnectedness of knowledge, values, and behavior"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Psychological Foundations:"
    p = tf.add_paragraph()
    p.text = "• Social-cognitive theory: values internalized through observation and reflection"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Moral development theory: ethical reasoning develops through stages"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Emotional intelligence: recognizing, understanding, managing emotions"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Values and Well-Being:"
    p = tf.add_paragraph()
    p.text = "• Values provide frameworks for decision-making"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Contribute to identity formation and sense of purpose"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Guide behavior promoting positive relationships"
    p.level = 1
    
    # SLIDE 7: BASAVANNA'S VACHANAS
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Traditional Indian Texts - Basavanna's Vachanas"
    tf = content.text_frame
    tf.text = "Basavanna's Vachanas (12th century CE):"
    p = tf.add_paragraph()
    p.text = "• Kannada philosophical poetry emphasizing social equality"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Key themes:"
    p = tf.add_paragraph()
    p.text = "  - Kayakave Kailasa (work is worship)"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Social equality and dignity"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Ethical conduct (achara)"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Self-awareness and introspection"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Service to others (seva)"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Educational Relevance:"
    p = tf.add_paragraph()
    p.text = "• Accessible language and concrete examples"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Emphasis on equality may support positive identity formation"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Limitations: Historical context may require adaptation; no empirical validation"
    p.level = 1
    
    # SLIDE 8: PANCHATANTRA
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Traditional Indian Texts - Panchatantra"
    tf = content.text_frame
    tf.text = "Panchatantra (3rd century BCE):"
    p = tf.add_paragraph()
    p.text = "• Classical collection of animal fables"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Originally composed as guide for princes on statecraft and life skills"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Key Themes:"
    p = tf.add_paragraph()
    p.text = "• Importance of wisdom and foresight"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Consequences of actions (karma)"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Cooperation and mutual benefit"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Ethical behavior and integrity"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Educational Relevance:"
    p = tf.add_paragraph()
    p.text = "• Narrative format makes abstract concepts concrete"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Allegorical nature allows discussion of sensitive topics"
    p.level = 1
    
    # SLIDE 9: HITOPADESHA
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Traditional Indian Texts - Hitopadesha"
    tf = content.text_frame
    tf.text = "Hitopadesha (800-950 CE):"
    p = tf.add_paragraph()
    p.text = "• Ancient Sanskrit text using animal fables"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Emphasizes practical wisdom and ethical decision-making"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Key Themes:"
    p = tf.add_paragraph()
    p.text = "• Prudence and careful consideration before action"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Foresight and understanding consequences"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Emotional restraint and self-control"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Wise counsel and learning from others"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Educational Relevance:"
    p = tf.add_paragraph()
    p.text = "• May support decision-making education"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Focus on emotional restraint aligns with regulation goals"
    p.level = 1
    
    # SLIDE 10: ANCIENT EDUCATION METHODS
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Ancient Education Methods"
    tf = content.text_frame
    tf.text = "Historical Value Transmission Methods:"
    p = tf.add_paragraph()
    p.text = "Panchatantra Stories:"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Explicitly designed as educational tool"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Pedagogical mechanisms: narrative engagement, allegorical distance"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "Scriptural Study:"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Process: recitation, commentary, reflection"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Functions: value transmission, cultural identity, moral reasoning"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "Pravachana (Oral Discourses):"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Primary method for transmitting values"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Characteristics: oral tradition, interactive dialogue, community learning"
    p.level = 2
    
    # SLIDE 11: METHODOLOGY
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Methodology"
    tf = content.text_frame
    tf.text = "Review Type: Narrative literature review (not systematic review)"
    p = tf.add_paragraph()
    p.text = "Time Period: 2015-2025"
    p = tf.add_paragraph()
    p.text = "Databases:"
    p = tf.add_paragraph()
    p.text = "• Google Scholar"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• PubMed"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Semantic Scholar"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Search Strategy:"
    p = tf.add_paragraph()
    p.text = "• Value-based education AND mental health"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Traditional Indian texts AND education"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Narrative approaches to moral/ethical education"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Analytical Approach: Scenario-based synthesis (analytical units, NOT case studies)"
    p = tf.add_paragraph()
    
    # SLIDE 12: RATIONALE FOR SCENARIO-BASED ANALYSIS
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Rationale for Scenario-Based Analysis"
    tf = content.text_frame
    tf.text = "Scenarios as Analytical Units:"
    p = tf.add_paragraph()
    p.text = "• Derived from common mental health challenges in literature"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Structured framework for conceptual exploration"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Each scenario examines: problem, traditional text, conceptual link, limitation"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Appropriate for Conceptual Research:"
    p = tf.add_paragraph()
    p.text = "• Allows exploration of potential applications"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Facilitates critical analysis of limitations"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Explicit Clarification:"
    p = tf.add_paragraph()
    p.text = "• Scenarios are NOT case studies"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Scenarios are NOT empirical evidence"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Scenarios are analytical tools for conceptual research"
    p.level = 1
    
    # SCENARIO SLIDES (12-21)
    scenarios = [
        {
            "num": 1,
            "title": "Academic Stress & Performance Anxiety",
            "problem": "High levels of stress, sleep disturbances, anxiety disorders in competitive academic environments (Pandey, 2019; Sharma, 2018)",
            "text": "Basavanna's Vachanas - emphasis on inner peace, detachment from external validation, intrinsic motivation",
            "link": "Philosophical shift from external metrics to internal growth might support stress reduction (theoretical connection)",
            "limitation": "No empirical studies validate effectiveness. Framework requires rigorous controlled testing. Abstract nature may not be accessible to all students."
        },
        {
            "num": 2,
            "title": "Peer Conflict & Aggression",
            "problem": "Peer conflicts and aggressive behaviors prevalent in schools, contributing to emotional distress and social isolation (Joshi & Mehta, 2021)",
            "text": "Hitopadesha - fables addressing conflict resolution, understanding perspectives, finding mutually beneficial solutions",
            "link": "Narrative-based discussions might provide frameworks for navigating interpersonal challenges. Allegorical format facilitates reflection.",
            "limitation": "Transfer of narrative lessons to real-world conflict resolution requires empirical validation. Cultural factors may limit generalizability."
        },
        {
            "num": 3,
            "title": "Low Self-Esteem & Identity Confusion",
            "problem": "Adolescent identity development issues, low self-worth associated with poor academic performance and social withdrawal (Singh & Verma, 2021)",
            "text": "Basavanna's Vachanas - emphasis on inherent worth of all individuals regardless of social status, promoting self-acceptance",
            "link": "Philosophical teachings about universal human value might conceptually support positive identity formation",
            "limitation": "No studies directly examine impact on self-esteem. Historical context may require adaptation. May not resonate across all cultural backgrounds."
        },
        {
            "num": 4,
            "title": "Bullying & Social Exclusion",
            "problem": "Bullying and social exclusion represent serious mental health threats, associated with depression, anxiety, suicidal ideation (Kumaran & Singh, 2019)",
            "text": "Panchatantra - stories about consequences of harmful behavior, importance of treating others with respect",
            "link": "Allegorical stories might provide safe spaces for discussing bullying. Narrative format could facilitate empathy development.",
            "limitation": "Narrative-based interventions for bullying require comprehensive empirical evaluation. Modern bullying involves digital platforms not addressed in traditional narratives."
        },
        {
            "num": 5,
            "title": "Anger & Emotional Dysregulation",
            "problem": "Emotional dysregulation, particularly anger management issues, impacts well-being and classroom dynamics. Linked to academic and social difficulties (Kumar, 2018)",
            "text": "Basavanna's Vachanas - emphasis on self-control, emotional mastery, understanding transient nature of emotions",
            "link": "Philosophical approaches to emotional awareness might conceptually support regulation skills. Aligns with cognitive-behavioral principles.",
            "limitation": "Abstract nature of philosophical teachings may not be accessible to all age groups. Direct translation to practical strategies requires empirical validation."
        },
        {
            "num": 6,
            "title": "Dishonesty & Ethical Drift",
            "problem": "Academic dishonesty and ethical lapses raise concerns about character development. Value-based interventions might support ethical decision-making (Pandey, 2019)",
            "text": "Hitopadesha and Panchatantra - stories illustrating consequences of dishonesty, value of integrity",
            "link": "Moral narratives might provide frameworks for ethical reasoning. Story format makes abstract concepts more concrete.",
            "limitation": "Relationship between narrative exposure and actual ethical behavior requires empirical investigation. Contemporary ethical challenges may not be addressed."
        },
        {
            "num": 7,
            "title": "Digital Addiction & Attention Fragmentation",
            "problem": "Digital device overuse leads to attention difficulties, sleep problems, social isolation. Increasing technology-related mental health issues (Nair & Sharma, 2020)",
            "text": "Basavanna's Vachanas - emphasis on mindfulness, present-moment awareness, concepts related to attention regulation",
            "link": "Philosophical teachings about focus and awareness might provide frameworks for understanding attention and distraction",
            "limitation": "Historical context predates digital technology entirely. Direct application requires significant adaptation. No empirical studies examine this connection."
        },
        {
            "num": 8,
            "title": "Peer Pressure & Impulsive Decision-Making",
            "problem": "Adolescent susceptibility to peer pressure and impulsive decision-making contributes to risky behaviors and negative mental health outcomes (Sharma & Rao, 2017)",
            "text": "Hitopadesha - stories about wise decision-making, considering consequences, resisting harmful influences",
            "link": "Stories about decision-making might provide frameworks for understanding choice, consequence, and peer influence",
            "limitation": "Narrative exposure alone may not translate to improved decision-making in real-world peer pressure situations. Comprehensive skill-building interventions likely necessary."
        },
        {
            "num": 9,
            "title": "Demotivation & Loss of Purpose",
            "problem": "Student demotivation and lack of purpose associated with poor academic performance and mental health issues. Meaning-making supports well-being (Narayan, 2015)",
            "text": "Basavanna's Vachanas - emphasis on finding purpose through service and ethical living",
            "link": "Philosophical discussions of purpose and meaning might conceptually support motivation and well-being",
            "limitation": "Abstract nature of purpose and meaning may be difficult for younger students to grasp. Spiritual context may require secular adaptation. Empirical validation lacking."
        },
        {
            "num": 10,
            "title": "Social Responsibility & Empathy Deficit",
            "problem": "Research indicates declining empathy levels and reduced engagement with social responsibility. Concerns about social cohesion and individual well-being (Desai & Patel, 2020)",
            "text": "All three sources emphasize social responsibility, compassion, concern for others' welfare",
            "link": "Narratives emphasizing empathy and social responsibility might provide frameworks for understanding interconnectedness and mutual care",
            "limitation": "Relationship between narrative exposure and actual empathetic behavior requires empirical investigation. Cultural and individual differences may affect narrative impact."
        }
    ]
    
    for scenario in scenarios:
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        content = slide.placeholders[1]
        
        title.text = f"Scenario {scenario['num']} - {scenario['title']}"
        tf = content.text_frame
        tf.text = "PROBLEM:"
        p = tf.add_paragraph()
        p.text = scenario['problem']
        p.level = 1
        p = tf.add_paragraph()
        p.text = "TRADITIONAL TEXT:"
        p = tf.add_paragraph()
        p.text = scenario['text']
        p.level = 1
        p = tf.add_paragraph()
        p.text = "CONCEPTUAL LINK:"
        p = tf.add_paragraph()
        p.text = scenario['link']
        p.level = 1
        p = tf.add_paragraph()
        p.text = "LIMITATION:"
        p = tf.add_paragraph()
        p.text = scenario['limitation']
        p.level = 1
    
    # SLIDE 23: SYNTHESIS - COMMON PATTERNS
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Synthesis - Common Patterns"
    tf = content.text_frame
    tf.text = "Common Patterns Identified:"
    p = tf.add_paragraph()
    p.text = "• Emotional Awareness and Regulation: Philosophical teachings about self-awareness might conceptually support emotional regulation"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Moral Reasoning and Ethical Decision-Making: Narrative approaches might provide frameworks for ethical reasoning"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Empathy and Social Understanding: Narratives might facilitate empathy development and perspective-taking"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Identity and Self-Concept: Philosophical teachings about inherent dignity might conceptually support positive identity formation"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Strengths of Value-Based Narratives:"
    p = tf.add_paragraph()
    p.text = "• Cultural relevance for students from Indian backgrounds"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Narrative engagement facilitates reflection"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Safe exploration through allegorical format"
    p.level = 1
    
    # SLIDE 24: SYNTHESIS - AREAS WHERE EVIDENCE IS WEAK
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Synthesis - Areas Where Evidence is Weak"
    tf = content.text_frame
    tf.text = "Critical Gaps:"
    p = tf.add_paragraph()
    p.text = "• Empirical Validation: No controlled studies examine traditional text-based interventions"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Mechanism Research: Limited understanding of how narratives might influence mental health"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Developmental Appropriateness: Insufficient research on age-appropriate applications"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Cultural Adaptation: Limited studies examining cross-cultural applicability"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Longitudinal Research: Absence of long-term follow-up studies"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Comparative Effectiveness: No studies comparing traditional narrative approaches to other interventions"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Implementation Research: Limited understanding of effective delivery methods"
    p.level = 1
    
    # SLIDE 25: CRITICAL LIMITATIONS
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Critical Limitations"
    tf = content.text_frame
    tf.text = "Absence of Empirical Validation:"
    p = tf.add_paragraph()
    p.text = "• NO evidence of intervention effectiveness"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• NO statistical improvements in mental health outcomes"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• NO behavioral changes from narrative exposure"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• All connections are conceptual and theoretical"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Other Limitations:"
    p = tf.add_paragraph()
    p.text = "• Subjectivity in interpretation"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Implementation challenges in schools"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Teacher training dependency"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Cultural generalization risks"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Methodological limitations (narrative review, not systematic)"
    p.level = 1
    
    # SLIDE 26: FUTURE RESEARCH DIRECTIONS
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Future Research Directions"
    tf = content.text_frame
    tf.text = "Research Priorities:"
    p = tf.add_paragraph()
    p.text = "Empirical Validation Models:"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Controlled intervention studies with randomized trials"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Quasi-experimental designs where randomization not feasible"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "Mixed-Method Studies:"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Quantitative measures of mental health outcomes"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Qualitative exploration of student experiences"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "Longitudinal Research:"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Sustained impacts over time"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Developmental trajectories"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "Additional Priorities: Teacher training, curriculum integration, assessment tools, comparative effectiveness"
    p.level = 1
    
    # SLIDE 27: CONCLUSION
    slide = prs.slides.add_slide(content_slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Conclusion"
    tf = content.text_frame
    tf.text = "Position Statement:"
    p = tf.add_paragraph()
    p.text = "This work is a CONCEPTUAL CONTRIBUTION:"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "  - Does NOT claim empirical effectiveness"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Does NOT recommend specific interventions"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Provides framework for understanding potential applications"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "  - Emphasizes critical need for rigorous empirical research"
    p.level = 2
    p = tf.add_paragraph()
    p.text = "Key Findings:"
    p = tf.add_paragraph()
    p.text = "• Traditional texts offer rich philosophical and narrative resources"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Scenario-based analysis reveals potential applications"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• Significant limitations and gaps identified"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "• All findings require empirical validation"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Future Directions: Careful empirical validation, cultural sensitivity, evidence-based implementation strategies"
    p = tf.add_paragraph()
    
    # Save presentation
    filename = "Value_Based_Education_Review_With_TOC.pptx"
    prs.save(filename)
    print(f"Presentation saved as {filename}")
    print(f"Total slides: {len(prs.slides)}")

if __name__ == "__main__":
    try:
        create_presentation()
        print("\nPowerPoint presentation created successfully!")
        print("\nNote: Review the presentation and adjust formatting as needed.")
    except ImportError:
        print("ERROR: python-pptx library not found.")
        print("Please install it using: pip install python-pptx")
    except Exception as e:
        print(f"Error creating presentation: {e}")

