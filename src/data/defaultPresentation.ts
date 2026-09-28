import type { Presentation } from "../types/presentation";

export const defaultPresentation: Presentation = {
  id: "sunday-service",
  title: "Sunday Service",
  slides: [
    {
      id: "welcome",
      type: "text",
      text: "Welcome to\nSunday Worship",
      background: "#18231f",
      textAlign: "center",
      fontSize: 64,
    },
    {
      id: "call-to-worship",
      type: "text",
      text: "Call to Worship\n\nLet us come before God with thankful hearts.",
      background: "#1d2524",
      textAlign: "center",
      fontSize: 52,
    },
    {
      id: "bible-reading",
      type: "text",
      text: "Bible Reading\n\nPlease open your Bible and follow along.",
      background: "#222522",
      textAlign: "center",
      fontSize: 50,
    },
    {
      id: "sermon",
      type: "text",
      text: "The Good News\n\nSunday Message",
      background: "#202622",
      textAlign: "center",
      fontSize: 58,
    },
    {
      id: "closing",
      type: "text",
      text: "Go in Peace\n\nHave a blessed week",
      background: "#1a2420",
      textAlign: "center",
      fontSize: 58,
    },
  ],
};