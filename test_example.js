/**
 * hwpx-generator 사용 예시
 * 실행: node test_example.js
 */
const { createGovHwpx, today } = require('./generate_hwpx_node');

// ── 서면보고 예시 ──
const report = {
  title: '디지털 전환 추진 현황 보고',
  doc_type: '서면보고',
  dept: '디지털정부혁신실',
  author: '홍길동',
  date: today(),
  sections: [
    {
      heading: '추진 배경',
      subsections: [
        {
          heading: '디지털 전환 필요성',
          paragraphs: [
            '코로나19 이후 비대면 행정서비스 수요가 급증하고 있다.',
            '주요국 대비 디지털 정부 경쟁력 강화가 시급한 상황이다.'
          ]
        }
      ],
      conclusions: ['디지털 전환 가속화를 위한 범정부 추진체계 마련 필요']
    },
    {
      heading: '주요 추진 내용',
      subsections: [
        {
          heading: 'AI 행정서비스 도입',
          paragraphs: ['민원 자동 분류, 문서 요약 등 AI 기반 서비스를 단계적으로 도입한다.'],
          table: {
            rows: [
              ['구분', '현재', '목표'],
              ['민원처리', '수동 분류', 'AI 자동분류'],
              ['문서작성', '개조식', 'AI 서술식']
            ]
          }
        }
      ]
    }
  ],
  contacts: [{ dept: '디지털정부혁신실', name: '홍길동', tel: '044-205-1234' }]
};

console.log('서면보고 생성 중...');
createGovHwpx(report, '서면보고_예시.hwpx')
  .then(path => console.log('✅ 서면보고 완료:', path))
  .catch(err => console.error('❌ 오류:', err));
