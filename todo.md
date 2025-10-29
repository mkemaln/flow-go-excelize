# Excel Arrow/Bent Connector Fix Plan

## Problem Analysis
- Excelize bent connectors cannot be flipped (X/Y direction)
- Line shapes always appear diagonally 
- Need proper bent connectors for flowchart arrows

## Solution Steps
- [ ] Analyze current arrow implementation
- [ ] Research excelize shape capabilities and limitations
- [ ] Design workaround for bent connectors using multiple shapes
- [ ] Implement custom positioning and rotation logic
- [ ] Test the solution with sample flowchart
- [ ] Verify arrow directions work correctly

## Technical Approach
1. Use combination of line segments to create L-shaped connectors
2. Implement proper rotation and positioning for diagonal lines
3. Create helper functions for different connector types
4. Test with existing query parameters

## Expected Outcome
- Proper bent connectors that can point in all directions
- Lines that don't appear diagonally when they should be straight
- Maintain compatibility with existing flowchart generation
