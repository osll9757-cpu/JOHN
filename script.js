// Presentation Navigation System
class Presentation {
    constructor() {
        this.slides = document.querySelectorAll('.slide');
        this.currentSlide = 0;
        this.totalSlides = this.slides.length;
        this.prevBtn = document.getElementById('prevBtn');
        this.nextBtn = document.getElementById('nextBtn');
        this.progressFill = document.getElementById('progressFill');
        
        this.init();
    }

    init() {
        // Set initial state
        this.updateSlide();
        
        // Event listeners
        this.prevBtn.addEventListener('click', () => this.previousSlide());
        this.nextBtn.addEventListener('click', () => this.nextSlide());
        
        // Keyboard navigation
        document.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowLeft') this.previousSlide();
            if (e.key === 'ArrowRight') this.nextSlide();
            if (e.key === 'Home') this.goToSlide(0);
            if (e.key === 'End') this.goToSlide(this.totalSlides - 1);
        });

        // Touch/swipe support
        let touchStartX = 0;
        let touchEndX = 0;

        document.addEventListener('touchstart', (e) => {
            touchStartX = e.changedTouches[0].screenX;
        });

        document.addEventListener('touchend', (e) => {
            touchEndX = e.changedTouches[0].screenX;
            this.handleSwipe();
        });

        const handleSwipe = () => {
            if (touchEndX < touchStartX - 50) this.nextSlide();
            if (touchEndX > touchStartX + 50) this.previousSlide();
        };

        this.handleSwipe = handleSwipe;
    }

    updateSlide() {
        // Remove all active/prev classes
        this.slides.forEach(slide => {
            slide.classList.remove('active', 'prev');
        });

        // Add active class to current slide
        this.slides[this.currentSlide].classList.add('active');

        // Add prev class to previous slides
        for (let i = 0; i < this.currentSlide; i++) {
            this.slides[i].classList.add('prev');
        }

        // Update button states
        this.prevBtn.disabled = this.currentSlide === 0;
        this.nextBtn.disabled = this.currentSlide === this.totalSlides - 1;

        // Update progress bar
        const progress = ((this.currentSlide + 1) / this.totalSlides) * 100;
        this.progressFill.style.width = `${progress}%`;

        // Announce slide change for accessibility
        this.announceSlide();
    }

    nextSlide() {
        if (this.currentSlide < this.totalSlides - 1) {
            this.currentSlide++;
            this.updateSlide();
            this.animateSlideContent();
        }
    }

    previousSlide() {
        if (this.currentSlide > 0) {
            this.currentSlide--;
            this.updateSlide();
            this.animateSlideContent();
        }
    }

    goToSlide(index) {
        if (index >= 0 && index < this.totalSlides) {
            this.currentSlide = index;
            this.updateSlide();
            this.animateSlideContent();
        }
    }

    animateSlideContent() {
        const currentSlideElement = this.slides[this.currentSlide];
        const content = currentSlideElement.querySelector('.slide-content');
        
        // Reset animation
        content.style.animation = 'none';
        setTimeout(() => {
            content.style.animation = 'fadeInUp 0.8s ease-out';
        }, 10);
    }

    announceSlide() {
        const slideNumber = this.currentSlide + 1;
        const announcement = `Slide ${slideNumber} of ${this.totalSlides}`;
        
        // Create or update aria-live region for screen readers
        let liveRegion = document.getElementById('slide-announcement');
        if (!liveRegion) {
            liveRegion = document.createElement('div');
            liveRegion.id = 'slide-announcement';
            liveRegion.setAttribute('aria-live', 'polite');
            liveRegion.setAttribute('aria-atomic', 'true');
            liveRegion.style.position = 'absolute';
            liveRegion.style.left = '-10000px';
            liveRegion.style.width = '1px';
            liveRegion.style.height = '1px';
            liveRegion.style.overflow = 'hidden';
            document.body.appendChild(liveRegion);
        }
        liveRegion.textContent = announcement;
    }
}

// Initialize presentation when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    const presentation = new Presentation();

    // Add smooth scroll behavior
    document.documentElement.style.scrollBehavior = 'smooth';

    // Prevent default arrow key scrolling
    window.addEventListener('keydown', (e) => {
        if(['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
            e.preventDefault();
        }
    });

    // Add loading animation
    document.body.style.opacity = '0';
    setTimeout(() => {
        document.body.style.transition = 'opacity 0.5s ease';
        document.body.style.opacity = '1';
    }, 100);

    // Console message
    console.log('%c🚀 Nuclear Bombs Presentation Loaded', 'color: #e94560; font-size: 16px; font-weight: bold;');
    console.log('%cKeyboard shortcuts:', 'color: #667eea; font-size: 14px; font-weight: bold;');
    console.log('← → : Navigate slides');
    console.log('Home : First slide');
    console.log('End : Last slide');
});

// Add particle effect on hover for cards (optional enhancement)
document.addEventListener('DOMContentLoaded', () => {
    const cards = document.querySelectorAll('.info-card, .comparison-card, .takeaway-item');
    
    cards.forEach(card => {
        card.addEventListener('mouseenter', function() {
            this.style.transition = 'all 0.3s cubic-bezier(0.68, -0.55, 0.265, 1.55)';
        });
    });
});

// Preload next slide for better performance
function preloadNextSlide(currentIndex) {
    const slides = document.querySelectorAll('.slide');
    if (currentIndex + 1 < slides.length) {
        const nextSlide = slides[currentIndex + 1];
        const images = nextSlide.querySelectorAll('img');
        images.forEach(img => {
            if (!img.complete) {
                const tempImg = new Image();
                tempImg.src = img.src;
            }
        });
    }
}

// Export for potential external use
if (typeof module !== 'undefined' && module.exports) {
    module.exports = Presentation;
}