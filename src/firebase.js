import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getStorage } from "firebase/storage";

// MISMO proyecto Firebase que las demás apps (costeador-de-recetas).
// No se crea proyecto nuevo. Config pública de cliente web.
const firebaseConfig = {
  apiKey: "AIzaSyCxb-8_XSZfl7BEtClFzp7C8pN4427G-DQ",
  authDomain: "costeador-de-recetas.firebaseapp.com",
  projectId: "costeador-de-recetas",
  storageBucket: "costeador-de-recetas.firebasestorage.app",
  messagingSenderId: "432314299669",
  appId: "1:432314299669:web:87802847a47c85dc33251c",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const storage = getStorage(app);
