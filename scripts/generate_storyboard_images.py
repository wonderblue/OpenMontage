import os
from PIL import Image, ImageDraw, ImageFont

def get_font(size=24, bold=False):
    # Try common font paths on macOS
    font_paths = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Microsoft/Arial Bold.ttf" if bold else "/Library/Fonts/Microsoft/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/Supplemental/Courier New.ttf"
    ]
    for path in font_paths:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except IOError:
                continue
    # Fallback to default
    return ImageFont.load_default()

def draw_slide(title, bullets, footer, filename, output_dir):
    # Dimensions for 16:9 aspect ratio (720p resolution)
    width = 1280
    height = 720
    
    # Premium Dark Slate Palette
    bg_color = (15, 23, 42)      # Deep slate #0F172A
    accent_color = (245, 158, 11)  # Warm amber #F59E0B
    text_color = (248, 250, 252)  # Clean white #F8FAFC
    footer_color = (148, 163, 184) # Muted gray #94A3B8
    
    # Create image
    img = Image.new("RGB", (width, height), bg_color)
    draw = ImageDraw.Draw(img)
    
    # Draw slide border/accents
    draw.rectangle([20, 20, width - 20, height - 20], outline=(51, 65, 85), width=2) # subtle card outline
    draw.rectangle([20, 20, 30, height - 20], fill=accent_color) # amber left-accent bar
    
    # Load fonts
    title_font = get_font(42, bold=True)
    body_font = get_font(26, bold=False)
    footer_font = get_font(18, bold=False)
    
    # Draw Title
    draw.text((60, 60), title, fill=accent_color, font=title_font)
    
    # Draw thin divider line
    draw.line([60, 130, width - 60, 130], fill=(51, 65, 85), width=2)
    
    # Draw Bullet Points
    y_pos = 180
    for bullet in bullets.split("\n"):
        if not bullet.strip():
            continue
        # Check if it's a sub-bullet or main bullet
        if bullet.strip().startswith("- ") or bullet.strip().startswith("* "):
            draw.text((100, y_pos), bullet.strip(), fill=text_color, font=body_font)
        else:
            draw.text((70, y_pos), bullet, fill=text_color, font=body_font)
        y_pos += 45
        
    # Draw Footer
    draw.text((width - 350, height - 60), footer, fill=footer_color, font=footer_font)
    
    # Save file
    os.makedirs(output_dir, exist_ok=True)
    save_path = os.path.join(output_dir, filename)
    img.save(save_path)
    print(f"Generated slide: {save_path}")

SLIDE_DATA = {
    "Video_1_Foundations": [
        ("1.1", "UNIT 1: Foundations of Marketing", 
         "• Importance & Scope of Marketing\n• The Marketing Ecosystem\n• Evolution of Marketing Thinking\n• The Concentric Marketing Environment\n• Customer Value & Industrial B2B Markets", 
         "Slide 1 | Introduction"),
        ("2.1", "Scope of Marketing - Part 1", 
         "• Market Research\n  - Guntur Chilli Market preparation for Rabi season\n• Strategy\n  - Target markets, value positioning, differentiation\n• Branding\n  - Guntur Sannam Chilli - GI Tag & reputation", 
         "Slide 1 | Scope & Importance"),
        ("2.2", "Scope of Marketing - Part 2", 
         "• Distribution\n  - AP Rythu Bazars (direct farmer-to-consumer sales)\n• Pricing\n  - Guntur Brodipet textiles 999/- psychology\n• Advertising\n  - AP Tourism beach & spiritual circuit promotions", 
         "Slide 1 | Scope & Importance"),
        ("2.3", "Scope of Marketing - Part 3", 
         "• Promotion\n  - Arundelpet restaurants local discount combos\n• Publicity\n  - Guntur Medical College national rankings news\n• Projects\n  - Amaravati capital city development masterplan", 
         "Slide 1 | Scope & Importance"),
        ("3.1", "Evolution of Concepts - Part 1", 
         "• Production Concept\n  - Mangalagiri handloom sarees volume & efficiency\n• Product Concept\n  - Kondapalli toy craftsmanship & heritage quality\n• Avoid 'Marketing Myopia'\n  - Never lose sight of underlying customer needs", 
         "Slides 2-3 | Marketing Concepts"),
        ("3.2", "Evolution of Concepts - Part 2", 
         "• Selling Concept\n  - Coercion & transactions vs. building long-term trust\n• Marketing Concept\n  - Sense & Respond (Guntur customized student courses)\n• Societal Concept\n  - Zero Budget Natural Farming (APCNF) sustainability", 
         "Slides 2-3 | Marketing Concepts"),
        ("4.1", "8 Tasks of Marketing Management", 
         "• Capture Insights & Connect with Customers\n• Build Strong Brands & Shape Offerings\n• Deliver & Communicate Value\n• Create Long-term Growth & Develop Strategies", 
         "Slide 4 | Management Tasks"),
        ("5.1", "The Concentric Marketing Environment", 
         "• Internal Environment: Organizational assets & culture\n• Micro Environment: Suppliers (farmers), competitors, customers\n• Macro Environment: GST policies, technological mandis, weather", 
         "Slide 5 | Marketing Environment"),
        ("6.1", "Customer Value Equation", 
         "• Value = Total Benefits - Total Costs\n• Benefits: Service quality, social recognition, emotions\n• Costs: Monetary price, time, physical effort, risk\n• Comparison: Guntur Main Bazar trust vs. Smart Bazaar", 
         "Slide 6 | Customer Value"),
        ("7.1", "Industrial Marketing (B2B)", 
         "• B2B: Managing relationships between organizations\n• Characteristics: Fewer but larger clients, derived demand\n• Examples: Vizag Steel raw contracts, Godrej Agrovet feed", 
         "Slide 7 | B2B Markets"),
        ("8.1", "Video 1 Summary & Recap", 
         "• Scope of Marketing & Evolutionary Concepts complete\n• 8 Tasks & Concentric Marketing Environment reviewed\n• Customer Value & B2B Industrial Marketing established", 
         "Slide 1 | Recap")
    ],
    "Video_2_Services": [
        ("1.1", "UNIT 1: Services Marketing", 
         "• Introduction to the Intangible Economy\n• Defining Services as Experiences & Transactions\n• Dynamic Trust Architecture (Uber Case Study)\n• 6 Characteristics of Services\n• The Extended 7Ps Mix for Services", 
         "Slide 1 | Services Introduction"),
        ("2.1", "Services Marketing & Uber Case Study", 
         "• Services Definition: Intangible transaction or performance\n• Uber Case: On-demand transportation technology\n• Trust Architecture: Double-sided ratings (4.9 vs 4.2)\n• Trust Guarantees: GPS tracking & live location sharing", 
         "Slide 9 | Uber Case Study"),
        ("3.1", "Service Characteristics - Part 1", 
         "• Intangibility: KIMS Hospital Guntur clean physical cues\n• Perishability: APSRTC bus empty seats lost revenue\n• Inseparability: Co-creation of education (Professor & Student)\n- Note: Customers co-create value in service delivery", 
         "Slide 11 | Service Characteristics"),
        ("3.2", "Service Characteristics - Part 2", 
         "• Variability: Guntur biryani spice level variance by chef\n• Changing Demand: Araku Valley winter peaks vs. summer valleys\n• Pricing: Perceived value & demand-based consulting rates", 
         "Slide 11 | Service Characteristics"),
        ("4.1", "7Ps of Services - Part 1", 
         "• P1: Product (Intangible core health care at Care Hospital)\n• P2: Price (Jio postpaid plans subscriptions, coaching fees)\n• P3: Place (Digital app distribution via Swiggy/Zomato)\n• P4: Promotion (Alumni success stories & placement records)", 
         "Slides 12-13 | Services Mix"),
        ("4.2", "7Ps of Services - Part 2", 
         "• P5: People (Staff emotional labor at Novotel Visakhapatnam)\n• P6: Process (Seamless digital citizen service at Mee Seva)\n• P7: Physical Evidence (Signature lobby scents, clean clinics)\n- Alignment: All 7Ps must reinforce the brand promise", 
         "Slides 12-13 | Services Mix"),
        ("5.1", "Video 2 Summary & Recap", 
         "• Services Marketing definition & characteristics verified\n• Uber Trust & Credibility model analyzed\n• 7Ps Services Mix framework established", 
         "Slide 1 | Services Recap")
    ],
    "Video_3_Global": [
        ("1.1", "UNIT 1: Global Marketing", 
         "• Introduction to International Trade Expansion\n• The Standardization vs. Adaptation Dilemma\n• Case Studies: Nike & Domino's Pizza\n• 8 Roadblocks & Challenges in Global Expansion", 
         "Slide 1 | Global Introduction"),
        ("2.1", "Global Marketing & The Core Dilemma", 
         "• Global Marketing: Expanding offerings into worldwide markets\n• Standardization: Uniform products (Apple iPhone economies of scale)\n• Adaptation: Customizing products to local cultures & laws\n• Glocalization: 'Think globally, act locally' (Starbucks Cardamom Chai)", 
         "Slide 14 | Global Dilemma"),
        ("3.1", "Global Giants: Nike & Domino's", 
         "• Nike: Sponsoring local sports (Cricket in India, Soccer in Europe)\n• Nike: Customizing cricket shoes with specific grass spike studs\n• Domino's Pizza: Szechuan toppings (Japan), Paneer Tikka (India)\n• Domino's Process: Custom delivery bikes for Mumbai traffic", 
         "Slides 15-16 | Case Studies"),
        ("4.1", "Global Challenges - Part 1", 
         "• 1. Cultural Differences: McDonald's McAloo Tikki (No Beef)\n• 2. Legal Compliance: GDPR privacy, Sweden TV ad bans on kids\n• 3. Economic: Single-use shampoo sachets in emerging markets\n• 4. Supply Chain: Temperature-controlled cold chain logistics", 
         "Slide 19 | Global Challenges"),
        ("4.2", "Global Challenges - Part 2", 
         "• 5. Political Risks: Tariffs, nationalization of assets\n• 6. Language Barriers: Slogan blunders (Nova Spanish: 'It doesn't go')\n• 7. Pricing: Parallel importing / textbook gray markets\n• 8. HR: US top-down leadership vs. Japanese Consensus (Ringi)", 
         "Slide 20 | Global Challenges"),
        ("5.1", "Video 3 Summary & Recap", 
         "• Global Marketing & Dilemmas addressed\n• Nike and Domino's localization case studies reviewed\n• 8 roadblocks & global challenges mapped", 
         "Slide 1 | Global Recap")
    ]
}

def generate_all_slides():
    master_dir = "/Users/rkanduk/Documents/GitHub/OpenMontage/professor_storyboard"
    
    # Render Video 1 Slides
    v1_dir = os.path.join(master_dir, "Video_1_Foundations")
    for scene, title, bullets, footer in SLIDE_DATA["Video_1_Foundations"]:
        filename = f"Video_1_Scene_{scene}_Slide.png"
        draw_slide(title, bullets, footer, filename, v1_dir)
        
    # Render Video 2 Slides
    v2_dir = os.path.join(master_dir, "Video_2_Services")
    for scene, title, bullets, footer in SLIDE_DATA["Video_2_Services"]:
        filename = f"Video_2_Scene_{scene}_Slide.png"
        draw_slide(title, bullets, footer, filename, v2_dir)
        
    # Render Video 3 Slides
    v3_dir = os.path.join(master_dir, "Video_3_Global")
    for scene, title, bullets, footer in SLIDE_DATA["Video_3_Global"]:
        filename = f"Video_3_Scene_{scene}_Slide.png"
        draw_slide(title, bullets, footer, filename, v3_dir)

if __name__ == "__main__":
    generate_all_slides()
