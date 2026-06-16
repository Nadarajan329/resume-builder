/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      fontFamily: {
        // We will define specific fonts, loaded dynamically in index.html
        cinzel: ['Cinzel', 'serif'],
        garamond: ['"EB Garamond"', 'serif'],
        inter: ['Inter', 'sans-serif'],
        cardo: ['Cardo', 'serif'],
        lora: ['Lora', 'serif'],
        sourceSans: ['"Source Sans 3"', 'sans-serif'],
      },
    },
  },
  plugins: [],
}
