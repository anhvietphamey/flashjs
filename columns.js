                                                          // 0-based; to be defined later

/** The first column is always text; others come from config. */
export const isTextColumn = (cfg             , col        ) =>
  col === 0 || (cfg.textColumns?.includes(col) ?? false);
