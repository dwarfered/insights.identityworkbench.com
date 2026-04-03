import {
  makeStyles,
  typographyStyles,
  BrandVariants,
  createLightTheme,
  Theme,
  tokens,
} from '@fluentui/react-components';

export const useQuickstartStyles = makeStyles({
  title: typographyStyles.title3,

  paragraph: {
    ...typographyStyles.body1,
    letterSpacing: '0.0675em',
    fontStyle: 'italic',
  },

  noBullets: {
    listStyleType: 'none',
    paddingLeft: 0,
  },

  toolbarNavButton: {
    backgroundColor: '#1c1c1c',
    color: tokens.colorBrandBackgroundInverted,
    ...typographyStyles.body1Strong,
    ':hover:active': {
      backgroundColor: '#333',
      color: '#fff',
    },
    ':hover': {
      backgroundColor: '#333',
      color: '#fff',
    },
  },

  card: {
    padding: '15px',
    marginTop: '24px',
    marginBottom: '24px',
    backgroundColor: 'rgb(255, 255, 255)',
    boxShadow:
      'rgba(0, 0, 0, 0.133) 0px 1.6px 3.6px 0px, rgba(0, 0, 0, 0.11) 0px 0.3px 0.9px 0px',
    transition: 'box-shadow 0.5s',
    ':hover': {
      boxShadow:
        'rgba(0, 0, 0, 0.18) 0px 3.2px 7.2px 0px, rgba(0, 0, 0, 0.14) 0px 0.6px 1.8px 0px',
    },
  },

  canvas: {
    paddingLeft: '8px',
    // paddingRight: '16px',
    paddingTop: '16px',
    paddingBottom: '16px',
  },

  canvasDense: {
    paddingLeft: '8px',
    paddingRight: '8px',
    paddingTop: '8px',
    paddingBottom: '8px',
    backgroundColor: 'rgb(250, 249, 248)',
  },
});

const ratingTheme: BrandVariants = {
  10: '#f56a00',
  20: '#e60001',
  30: '#990100',
  40: '#193253',
  50: '#1B3F6A',
  60: '#1B4C82',
  70: '#18599B',
  80: '#1267B4',
  90: '#3174C2',
  100: '#4F82C8',
  110: '#6790CF',
  120: '#7D9ED5',
  130: '#92ACDC',
  140: '#A6BAE2',
  150: '#BAC9E9',
  160: '#CDD8EF',
};

export const lowRatingTheme: Theme = {
  ...createLightTheme(ratingTheme),
};
lowRatingTheme.colorBrandForeground1 = ratingTheme[10];

export const mediumRatingTheme: Theme = {
  ...createLightTheme(ratingTheme),
};
mediumRatingTheme.colorBrandForeground1 = ratingTheme[20];

export const highRatingTheme: Theme = {
  ...createLightTheme(ratingTheme),
};
highRatingTheme.colorBrandForeground1 = ratingTheme[30];
