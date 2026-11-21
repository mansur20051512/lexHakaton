from django.db import models

class LegalCase(models.Model):
    """Fuqarolarning tahlil qilingan huquqiy murojaatlari va yaratilgan hujjatlar tarixi"""
    
    INTENT_CHOICES = [
        ('mediation', 'Mediasiya (Sudgacha tinch hal qilish)'),
        ('appeal', 'Rasmiy Murojaat (Davlat organlariga)'),
        ('notary', 'Notariat (Xizmatlar va Bojlar)'),
    ]
    
    LANG_CHOICES = [
        ('uz_latn', 'O‘zbekcha (Lotin)'),
        ('uz_cyrl', 'Ўзбекча (Кирилл)'),
        ('ru', 'Русский'),
    ]

    # Asosiy murojaat
    query_text = models.TextField(verbose_name="Fuqaro murojaati matni")
    detected_language = models.CharField(max_length=20, choices=LANG_CHOICES, default='uz_latn', verbose_name="Aniqlangan til")
    primary_intent = models.CharField(max_length=30, choices=INTENT_CHOICES, default='mediation', verbose_name="Asosiy yo‘nalish")
    secondary_intent = models.CharField(max_length=30, blank=True, null=True, verbose_name="Qo‘shimcha yo‘nalish")
    urgency_level = models.CharField(max_length=30, default='o‘rtacha', verbose_name="Shoshilinchlik darajasi")
    case_category = models.CharField(max_length=255, default='Fuqarolik-huquqiy nizo', verbose_name="Ish toifasi")

    # Yuridik xulosa va maslahat
    headline = models.CharField(max_length=500, verbose_name="Qisqa xulosa")
    detailed_advice = models.TextField(verbose_name="Batafsil yuridik maslahat")

    # Taqqoslash (Sud vs Mediasiya)
    court_duration = models.CharField(max_length=100, default='2 – 4 oy', verbose_name="Sud muddati")
    court_cost_estimate = models.CharField(max_length=255, default='4% davlat boji + advokat', verbose_name="Sud xarajati")
    mediation_duration = models.CharField(max_length=100, default='3 – 7 kun', verbose_name="Mediasiya muddati")
    mediation_cost_estimate = models.CharField(max_length=255, default='Minimal / kelishuv asosida', verbose_name="Mediasiya xarajati")
    peaceful_resolution_chance = models.PositiveIntegerField(default=85, verbose_name="Murosa ehtimoli (%)")

    # Rasmiy hujjat
    document_title = models.CharField(max_length=255, default='TALABNOMA', verbose_name="Hujjat nomi")
    document_type = models.CharField(max_length=100, default='TALABNOMA', verbose_name="Hujjat turi")
    recipient_title = models.CharField(max_length=500, blank=True, default='', verbose_name="Kimga (Adresat)")
    sender_placeholder = models.CharField(max_length=500, blank=True, default='', verbose_name="Kimdan (Arizachi)")
    document_body = models.TextField(verbose_name="Hujjatning to‘liq matni")
    attachments_json = models.JSONField(default=list, blank=True, verbose_name="Ilovalar ro‘yxati")

    # Xom tahlil ma'lumotlari
    raw_response_json = models.JSONField(default=dict, blank=True, verbose_name="AI to‘liq JSON javobi")
    docx_file = models.FileField(upload_to='generated_docs/', blank=True, null=True, verbose_name="Word (.docx) fayli")
    
    # Texnik metrikalar
    execution_time_seconds = models.FloatField(default=0.0, verbose_name="Bajarilish vaqti (soniya)")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Yaratilgan vaqti")

    class Meta:
        verbose_name = "Huquqiy Murojaat"
        verbose_name_plural = "Huquqiy Murojaatlar"
        ordering = ['-created_at']

    def __str__(self):
        return f"[{self.get_primary_intent_display()}] - {self.headline[:60]} ({self.created_at.strftime('%Y-%m-%d %H:%M')})"
