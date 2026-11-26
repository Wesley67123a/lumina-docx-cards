import { buildDocRow } from './docxParser.js';

export const SAMPLE_DOC_DATA = {
  title: 'HKDSE BIOLOGY — Chapter 8 & 17 Q&A Notes',
  fileName: 'HKDSE_Bio_Ch2_to_Ch27_QA_Notes-v48_Goodnotes.docx',
  driveFileId: null,
  rows: [
    buildDocRow({
      id: 'row-bio-1',
      index: 1,
      section: 'Chapter 8: Transport in Humans',
      questionLines: [
        'Describe the pressure changes in the blood as it flows through the circulatory system from the aorta to the veins. (3 marks)',
        '*OCR GCE (AS) Cells, Exchange and Transport May 2012 Q3b',
      ],
      answerLines: [
        'The blood pressure in aorta changes periodically. (1)',
        'It drops significantly in the arteries. (1)',
        'The blood pressure is constant in veins. (1)',
      ],
      pageBreakAfter: false,
    }),
    buildDocRow({
      id: 'row-bio-2',
      index: 2,
      section: 'Chapter 8: Transport in Humans',
      questionLines: [
        'Explain why it is important that the pressure changes as blood flows from the aorta to the capillaries. (2 marks)',
        '*OCR GCE (AS) Cells, Exchange and Transport May 2012 Q3c(ii)',
      ],
      answerLines: [
        'The capillary wall is one-cell thick. (1)',
        'High pressure would damage the capillary wall. (1)',
      ],
      pageBreakAfter: false,
    }),
    buildDocRow({
      id: 'row-bio-3',
      index: 3,
      section: 'Chapter 8: Transport in Humans',
      questionLines: [
        'How can damaged valves in veins lead to varicose veins? (2 marks)',
        '*Oxford Ch.8 QB (p. 8-95)',
      ],
      answerLines: [
        'The damaged valves in veins allow the backflow of blood. (1)',
        'Therefore blood accumulates in veins, causing the veins to get larger and bulge on the skin surface. (1)',
      ],
      pageBreakAfter: false,
    }),
    buildDocRow({
      id: 'row-bio-4',
      index: 4,
      section: 'Chapter 8: Transport in Humans',
      questionLines: [
        'Why are people who need to stand for a prolonged period of time at higher risk of having varicose veins in their lower limbs? (1 mark)',
        '*Oxford Ch.8 QB (p. 8-95)',
      ],
      answerLines: [
        'Standing for a long period of time increases the blood volume in the veins in the lower limbs due to gravity. Thus, blood accumulates in the veins more easily. (1)',
      ],
      pageBreakAfter: true,
    }),
    buildDocRow({
      id: 'row-bio-5',
      index: 5,
      section: 'Chapter 17: Coordination in Humans',
      questionLines: [
        'Many mitochondria can be found in the axon endings of the motor neurones. State the roles of mitochondria in the transmission of nerve impulses across the neuromuscular junctions. (4 marks)',
        '*HKDSE Past Paper Practice',
      ],
      answerLines: [
        'Mitochondria are the main site of respiration. (1)',
        'They release energy from the food for various activities in the neuromuscular junctions: (1)',
        'Energy is used in the synthesis of neurotransmitter. / Energy is used in the formation of vesicles. / Energy is used in the movement of vesicles. / Energy is used in exocytosis / the release of neurotransmitter into the cleft. (any 2, 1m × 2)',
      ],
      pageBreakAfter: true,
    }),
    buildDocRow({
      id: 'row-bio-6',
      index: 6,
      section: 'Chapter 17: Coordination in Humans',
      questionLines: [
        'The bacterium Clostridium botulinum releases a toxin which can bind to presynaptic membrane P and can cause death as it can lead to paralysis of the human breathing system. Explain how the binding of the toxin released by Clostridium botulinum to membrane P leads to the paralysis of the human breathing system. (3 marks)',
        '*HKDSE Bio Ch.17 QB',
      ],
      answerLines: [
        'The toxin binds to membrane P and blocks the release of neurotransmitter. (1)',
        'When a nerve impulse reaches the ending of the motor neurone, no neurotransmitter is released and thus muscle fibres are not stimulated to contract. (1)',
        'If the intercostal muscles cannot contract, the breathing process cannot be carried out. (1)',
      ],
      pageBreakAfter: false,
    }),
    buildDocRow({
      id: 'row-bio-7',
      index: 7,
      section: 'Chapter 18: Movement in Humans',
      questionLines: [
        'The end surfaces of the upper leg bone are covered with cartilage. Give two functions of this cartilage at the hip joint. (2 marks)',
        '*HKDSE Bio Ch.18 QB',
      ],
      answerLines: [
        'It absorbs shock during movement at the hip joint and the knee joint. (1)',
        'It reduces friction between the hip bone and the upper leg bone, and between the leg bones during movement. (1)',
      ],
      pageBreakAfter: false,
    }),
    buildDocRow({
      id: 'row-bio-8',
      index: 8,
      section: 'Chapter 8: Transport in Humans',
      questionLines: [
        'How can compression stockings treat varicose veins? (3 marks)',
        '*Oxford Ch.8 QB (p. 8-95)',
      ],
      answerLines: [
        'Compression stockings exert an external pressure on the skeletal muscles in the lower limbs. (1)',
        'This promotes the flow of blood back to the heart. (1)',
        'Therefore, accumulation of blood in the veins is prevented. (1)',
      ],
      pageBreakAfter: false,
    }),
  ],
};

/**
 * Initial pre-added Flashcards (so Day 1 / Day 2 tracking, Overview Bar Chart,
 * and Docs "✓ 已加入" checkmarks are immediately visible on first open!)
 */
export function buildInitialSeededFlashcards(startDateStr) {
  const d1 = startDateStr;
  const d2 = addDaysISO(startDateStr, 1);
  return {
    'row-bio-1': {
      rowId: 'row-bio-1',
      addedDay: 1,
      addedDate: d1,
      mastered: true,
      bookmarked: false,
      checkedPoints: [0, 1, 2],
    },
    'row-bio-2': {
      rowId: 'row-bio-2',
      addedDay: 1,
      addedDate: d1,
      mastered: true,
      bookmarked: true,
      checkedPoints: [0, 1],
    },
    'row-bio-5': {
      rowId: 'row-bio-5',
      addedDay: 2,
      addedDate: d2,
      mastered: false,
      bookmarked: true,
      checkedPoints: [0],
    },
    'row-bio-6': {
      rowId: 'row-bio-6',
      addedDay: 2,
      addedDate: d2,
      mastered: false,
      bookmarked: false,
      checkedPoints: [],
    },
  };
}

export function addDaysISO(isoDate, deltaDays) {
  const d = new Date(`${isoDate}T00:00:00`);
  d.setDate(d.getDate() + deltaDays);
  return toLocalISO(d);
}

export function toLocalISO(date = new Date()) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}
