const graphBase = 'https://graph.microsoft.com/v1.0';

export const msGraphEndpoints = {
  graphMe: `${graphBase}/me`,
  graphMePhoto: `${graphBase}/me/photos/48x48/$value`,
  graphUsers: `${graphBase}/users`,
  groups: `${graphBase}/groups`,
  subscribedSkus: `${graphBase}/subscribedSkus`,
  organization: `${graphBase}/organization`,
};
