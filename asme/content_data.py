"""Static site copy: mission, highlights, showcase and executive profile overrides.

Edited by hand when the exec board changes; everything else on the public site
comes from the database.
"""

FRONT_CLUB_MISSION = (
    "ASME at Iowa is a hands-on engineering organization where members design, build, "
    "test, and iterate real mechanical systems while developing leadership and teamwork."
)

FRONT_CLUB_HIGHLIGHTS = [
    {
        "label": "Design + Fabrication",
        "text": "Members move from CAD to manufacturing and validation in real build cycles.",
    },
    {
        "label": "Technical Leadership",
        "text": "Student leads coordinate subsystems, reviews, and project execution timelines.",
    },
    {
        "label": "Industry Readiness",
        "text": "Project workflows mirror engineering practice: requirements, testing, and documentation.",
    },
]

FRONT_PROJECT_SHOWCASE = [
    {
        "team": "Baja Team",
        "name": "Off-Road Vehicle Program",
        "summary": "End-to-end student-built vehicle development with subsystem integration and track testing.",
        "status": "Active Build Season",
    },
    {
        "team": "Formula Team",
        "name": "Formula Design Initiative",
        "summary": "Performance-focused design loops covering chassis, powertrain, controls, and test data analysis.",
        "status": "Prototype + Validation",
    },
    {
        "team": "Design Team",
        "name": "Crater Crusher Platform",
        "summary": "Mission-driven mechanical system development for robust field operation and reliability.",
        "status": "Iteration + Review",
    },
]

FRONT_ABOUT_ASME_FACTS = [
    "ASME is a not-for-profit membership organization focused on collaboration, knowledge sharing, and career development across engineering disciplines.",
    "ASME was founded in 1880 and now includes more than 100,000 members across 140+ countries.",
    "About 32,000 ASME members are students.",
]

FRONT_UIOWA_CURRENT_PROJECTS = [
    {
        "name": "Design Build Fly",
        "summary": (
            "Teams design, fabricate, and demonstrate an unmanned electric radio-controlled aircraft "
            "to meet a defined mission profile."
        ),
    },
    {
        "name": "Additive Manufacturing Mars Rover (R.O.V.E.R.)",
        "summary": (
            "Students use additive manufacturing and iterative design to build an unmanned vehicle "
            "that gathers and deposits resources in an extraterrestrial-style environment."
        ),
    },
    {
        "name": "Automated Garbage Truck",
        "summary": (
            "Student design teams build and test a waste collection system that navigates a model city, "
            "sorts waste streams, and delivers them to the correct destination."
        ),
    },
]

EXEC_PROFILE_OVERRIDES = {
    "email:brayden-nagra@uiowa.edu": {
        "title": "President",
        "headshot_url": "/static/images/executive/brayden_nagra.jpg",
        "message": (
            "As President, I set the overall direction for ASME at Iowa, coordinate the executive board, "
            "and represent our chapter to the College of Engineering, sponsors, and national ASME.\n"
            "This semester my main focus is building sustainable systems for funding, mentorship, and project "
            "management so we can support 150-200+ active members and multiple national competition teams.\n"
            "Members can come to me if they want to get plugged into a project team, start a new initiative, "
            "or talk about sponsorships, leadership opportunities, or long-term plans for the club.\n"
            "brayden-nagra@uiowa.edu, 806-577-2216"
        ),
    },
    "name:braydennagra": {
        "title": "President",
        "headshot_url": "/static/images/executive/brayden_nagra.jpg",
        "message": (
            "As President, I set the overall direction for ASME at Iowa, coordinate the executive board, "
            "and represent our chapter to the College of Engineering, sponsors, and national ASME.\n"
            "This semester my main focus is building sustainable systems for funding, mentorship, and project "
            "management so we can support 150-200+ active members and multiple national competition teams.\n"
            "Members can come to me if they want to get plugged into a project team, start a new initiative, "
            "or talk about sponsorships, leadership opportunities, or long-term plans for the club.\n"
            "brayden-nagra@uiowa.edu, 806-577-2216"
        ),
    },
    "email:jaetnyre@uiowa.edu": {
        "title": "Executive Coordinator",
        "headshot_url": "/static/images/executive/jet_etnyre.jpg",
        "message": (
            "As Executive Coordinator, I handle a lot of scheduling work across ASME operations.\n"
            "I use ASTRA Schedule Viewer, submit room request forms, and post events on Engage once reservations are confirmed.\n"
            "I also support external events outside GBMs. I created the Dallas trip itinerary, worked with Eddy on hotel planning, helped set up volleyball events last semester, and I am helping run Makeathon this semester.\n"
            "My personal goal this semester is to win an intramural championship, and for ASME I am focused on making the Dallas trip run smoothly.\n"
            "Members can get involved by attending GBMs and events. I would love ideas for next-year events outside GBMs (socials, games, sports, etc).\n"
            "jaetnyre@uiowa.edu"
        ),
    },
    "name:jetetnyre": {
        "title": "Executive Coordinator",
        "headshot_url": "/static/images/executive/jet_etnyre.jpg",
        "message": (
            "As Executive Coordinator, I handle a lot of scheduling work across ASME operations.\n"
            "I use ASTRA Schedule Viewer, submit room request forms, and post events on Engage once reservations are confirmed.\n"
            "I also support external events outside GBMs. I created the Dallas trip itinerary, worked with Eddy on hotel planning, helped set up volleyball events last semester, and I am helping run Makeathon this semester.\n"
            "My personal goal this semester is to win an intramural championship, and for ASME I am focused on making the Dallas trip run smoothly.\n"
            "Members can get involved by attending GBMs and events. I would love ideas for next-year events outside GBMs (socials, games, sports, etc).\n"
            "jaetnyre@uiowa.edu"
        ),
    },
    "email:arkaiser@uiowa.edu": {
        "name": "Adelai Kaiser",
        "title": "Treasurer",
        "headshot_url": "/static/images/executive/adelai_kaiser.jpg",
        "message": (
            "The treasurer role handles the finances of the club, coordinating with SOBO, and keeping track of the club's expenses.\n"
            "My goal this semester is to learn more about PCBs.\n"
            "Members can get involved by coming to the club's general meetings which are listed on Engage.\n"
            "arkaiser@uiowa.edu"
        ),
    },
    "name:akaiser": {
        "name": "Adelai Kaiser",
        "title": "Treasurer",
        "headshot_url": "/static/images/executive/adelai_kaiser.jpg",
        "message": (
            "The treasurer role handles the finances of the club, coordinating with SOBO, and keeping track of the club's expenses.\n"
            "My goal this semester is to learn more about PCBs.\n"
            "Members can get involved by coming to the club's general meetings which are listed on Engage.\n"
            "arkaiser@uiowa.edu"
        ),
    },
    "name:adelaikaiser": {
        "name": "Adelai Kaiser",
        "title": "Treasurer",
        "headshot_url": "/static/images/executive/adelai_kaiser.jpg",
    },
}

MANUAL_EXECUTIVE_PROFILES = [
    {
        "name": "Brayden Nagra",
        "title": "President",
        "headshot_url": "/static/images/executive/brayden_nagra.jpg",
        "message": (
            "As President, I set the overall direction for ASME at Iowa, coordinate the exec board, "
            "and represent our chapter to the College of Engineering, sponsors, and national ASME.\n"
            "This semester my main focus is building sustainable systems for funding, mentorship, and "
            "project management so we can support 150-200+ active members and multiple national competition teams.\n"
            "Members can come to me if they want to get plugged into a project team, start a new initiative, "
            "or talk about sponsorships, leadership opportunities, or long-term plans for the club.\n"
            "brayden-nagra@uiowa.edu, 806-577-2216"
        ),
    },
    {
        "name": "Paul Conover",
        "title": "Vice President",
        "headshot_url": "/static/images/executive/paul_conover.jpg",
        "message": (
            "As Vice President, I reach out to industry partners, assist with club direction, and oversee project teams.\n"
            "My goal for ASME is to be a vessel for members to progress in their education and careers. "
            "I want to expand ASME at Iowa to create more technically ambitious engineers who have the drive required to succeed in industry.\n"
            "Companies looking to get involved with the ASME club can reach out to me, as well as any students who want to join.\n"
            "pconover@uiowa.edu, 712-346-8176"
        ),
    },
    {
        "name": "Adelai Kaiser",
        "title": "Treasurer",
        "headshot_url": "/static/images/executive/adelai_kaiser.jpg",
        "message": (
            "The treasurer role handles the finances of the club, coordinating with SOBO, and keeping track of the club's expenses.\n"
            "My goal this semester is to learn more about PCBs.\n"
            "Members can get involved by coming to the club's general meetings which are listed on Engage.\n"
            "arkaiser@uiowa.edu"
        ),
    },
    {
        "name": "Jet Etnyre",
        "title": "Executive Coordinator",
        "headshot_url": "/static/images/executive/jet_etnyre.jpg",
        "message": (
            "As exec coordinator, it's a lot of scheduling — going onto ASTRA schedule viewer, filling out the room request form, "
            "and submitting the event on Engage once I get the room reservation back. "
            "I also support external events outside GBMs. I created the Dallas trip itinerary, worked with Eddy on hotel planning, helped set up volleyball events last semester, and I am helping run Makeathon this semester.\n"
            "One goal for this semester personally is to win an intramural championship, but ASME-wise I am just focused on having the Dallas trip run smoothly.\n"
            "Members can get involved by attending GBMs and events. I would love ideas for next-year events outside GBMs (socials, games, sports, etc).\n"
            "jaetnyre@uiowa.edu"
        ),
    },
    {
        "name": "Dawson Fish",
        "title": "Make-A-Thon Coordinator",
        "headshot_url": "/static/images/executive/dawson_fish.jpg",
        "message": (
            "My role as Make-A-Thon Coordinator is to facilitate and plan this coming year's Make-A-Thon. "
            "This involves selecting a time, a prompt, the food that will be provided, and the judges for the event.\n"
            "My goal for next semester is to shift the mindsets of our club members to emphasize career growth and hands-on experience. "
            "I want all who are associated with ASME to fully buy into the idea that heavy involvement that develops both technical and soft skills is crucial to a successful career.\n"
            "Members can get involved through the Make-A-Thon — it's a great chance to experience the real engineering design process. Reach out about Make-A-Thon, internships/co-op, class recommendations, or just to chat.\n"
            "dmfish@uiowa.edu"
        ),
    },
    {
        "name": "Nathan Fish",
        "title": "Shop Manager",
        "headshot_url": "/static/images/executive/nathan_fish.jpg",
        "message": (
            "The Shop Manager protects ASME's physical assets, especially in G440 and shared lab spaces. "
            "This role is critical for cost control, safety, and professionalism.\n"
            "My goal for next year is to maintain a clean, organized, and efficient workspace so members can easily find and use tools. "
            "I want to reinforce good habits to keep the space organized and reduce downtime so people can focus on building rather than searching for tools.\n"
            "Members can get involved by helping keep G440 clean and treating the shared tools with care — and come to me with shop questions or training requests.\n"
            "nfsh@uiowa.edu"
        ),
    },
    {
        "name": "Heriberto Salgado",
        "title": "Event Coordinator",
        "headshot_url": "/static/images/executive/heriberto_salgado.jpg",
        "message": (
            "I serve as the Event Coordinator for ASME. In this role, I help organize events such as the ACE Mentor Program visit and the annual IAM3D event.\n"
            "My goal is to create more engaging opportunities for members to connect, learn, and get involved this semester.\n"
            "Members are always welcome to come to me for CAD modeling advice or guidance with mechanical engineering classes. The best way to reach me is through Teams or by email.\n"
            "hsalgado@uiowa.edu"
        ),
    },
    {
        "name": "Henry Strauss",
        "title": "Marketing Director",
        "headshot_url": "/static/images/executive/henry_strauss.jpg",
        "message": (
            "I lead marketing and media for ASME — content, photography, video, and the campaigns that grow interest in the club.\n"
            "My big focus this year is to bring innovation to ASME through new marketing campaigns to grow interest both internally and externally for all ASME-related events. "
            "I believe high-quality, eye-catching media is key to professionalism and growth.\n"
            "Members can get involved by helping shoot/edit content for events, project teams, and socials — reach out if you want to be on the media team.\n"
            "hjstrauss@uiowa.edu"
        ),
    },
]

